import requests


def send(mess):
    chat_id = '-590657677'
    token = '5481283557:AAFxR7VffxwFkh7tDx6gLgx8NAKSlM6B6hQ'
    # msg = "Send text with photo 😉"
    telegram_msg = requests.get(f'https://api.telegram.org/bot{token}/sendMessage?chat_id={chat_id}&text={mess}')
    # print(telegram_msg)
    # print(telegram_msg.content)

