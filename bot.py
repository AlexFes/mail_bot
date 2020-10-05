import logging
import os
from telegram import ParseMode
from telegram.constants import MAX_MESSAGE_LENGTH
from telegram.ext import (Updater, CommandHandler)
from utils.client import EmailClient


logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s:%(lineno)d'
                           ' - %(message)s', filename='../mailbot.log',
                    level=logging.INFO)
logger = logging.getLogger(__name__)

bot_token = os.environ['TELEGRAM_TOKEN']

def handle_large_text(text):
    while text:
        if len(text) < MAX_MESSAGE_LENGTH:
            yield text
            text = None
        else:
            out = text[:MAX_MESSAGE_LENGTH]
            yield out
            text = text.lstrip(out)

def error(bot, update, _error):
    logger.warning('Update "%s" caused error "%s"', update, _error)

def start_callback(bot, update):
    msg = "Use /help to get help"
    update.message.reply_text(msg)

def _help(bot, update):
    help_str = "*Mailbox Setting*: \n" \
               "/setting 123456@example.com yourpassword \n" \
               "*Get 'n' latest messages (3 by default):* \n" \
               "/get n\n" \
               "*Reply the last message*: \n" \
               "/reply Enter your text here \n" \
               "*Reply to 'email'*: \n" \
               "/replyto email Enter your text here"
    bot.send_message(update.message.chat_id,
                    parse_mode=ParseMode.MARKDOWN,
                    text=help_str)

def setting_email(bot, update, args, job_queue, chat_data):
    global email_addr, email_passwd, inbox_num
    chat_id = update.message.chat_id
    email_addr = args[0]
    email_passwd = args[1]
    logger.info("received setting_email command.")
    update.message.reply_text("Configure email success!")
    with EmailClient(email_addr, email_passwd) as client:
        inbox_num = client.get_mails_count()
    job = job_queue.run_repeating(periodic_task, 5, context=chat_id)
    chat_data['job'] = job
    logger.info("periodic task scheduled.")


def periodic_task(bot, job):
    global inbox_num
    logger.info("entering periodic task.")
    with EmailClient(email_addr, email_passwd) as client:
        new_inbox_num = client.get_mails_count()
        if new_inbox_num > inbox_num:
            mail = client.get_mail_by_index(1)
            content = mail.__repr__()
            for text in handle_large_text(content):
                bot.send_message(job.context,
                                text=text)
            inbox_num = new_inbox_num

def get_email(bot, update, args):
    count = 3
    if (len(args)):
        count = int(args[0])
    logger.info("received get command.")
    with EmailClient(email_addr, email_passwd) as client:
        for i in range(1, count + 1):
            mail = client.get_mail_by_index(i)
            content = mail.__repr__()
            for text in handle_large_text(content):
                bot.send_message(update.message.chat_id,
                                 text=text)

def reply(bot, update, args):
    logger.info("recieved reply command.")
    try:
        with EmailClient(email_addr, email_passwd) as client:
            last_mail = client.get_mail_by_index(1)

            to, subject = last_mail.get_reply_data()
            text = " ".join(args)
            logger.info("to = {} subject = {} text = {}".format(to, subject, text))
            client.send_mail(to, subject, text)
            bot.send_message(update.message.chat_id, text="Message sent")
    except:
        bot.send_message(update.message.chat_id, text="Reply failed")

def replyto(bot, update, args):
    logger.info("recieved replyto command.")
    try:
        with EmailClient(email_addr, email_passwd) as client:
            to, subject = (args[0], "Reply")
            text_list = args[1:]
            text = " ".join(text_list)
            logger.info("to = {} subject = {} text = {}".format(to, subject, text))
            client.send_mail(to, subject, text)
            bot.send_message(update.message.chat_id, text="Message sent")
    except:
        bot.send_message(update.message.chat_id, text="Reply failed")

def main():
    updater = Updater(token=bot_token)
    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start_callback))
    dp.add_handler(CommandHandler("help", _help))
    dp.add_handler(CommandHandler("setting", setting_email, pass_args=True,
                                  pass_job_queue=True, pass_chat_data=True))

    dp.add_handler(CommandHandler("get", get_email, pass_args=True))
    dp.add_handler(CommandHandler("reply", reply, pass_args=True))
    dp.add_handler(CommandHandler("replyto", replyto, pass_args=True))
    dp.add_error_handler(error)

    updater.start_polling()
    updater.idle()


if __name__ == '__main__':
    main()
