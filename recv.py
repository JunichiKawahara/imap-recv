#! /usr/local/bin/python3

import email
import ssl
import imaplib
import os
from os.path import join, dirname
from dotenv import load_dotenv
from email.header import decode_header, make_header

"""
recv.py
IMAPサーバからメッセージを取得するテストスクリプト
Copyright (C) 2019 J.Kawahara
2019.3.4 J.Kawahara 新規作成
"""

"""
以下の準備が必要！
$ pip install python-dotenv
"""

# 環境変数を取得する
dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

MAX_FETCH_NUM = int(os.environ.get("MAX_FETCH_NUM"))
MAIL_SERVER = os.environ.get("MAIL_SERVER")
MAIL_USER = os.environ.get("MAIL_USER")
MAIL_PASS = os.environ.get("MAIL_PASS")
ENC_TYPE = os.environ.get("ENC_TYPE")
PORT = int(os.environ.get("PORT"))


def receive_mail():
    """メールを受信する
    """

    nego_combo = (ENC_TYPE, PORT)

    imapclient = None
    if nego_combo[0] == "no-encrypt":
        imapclient = imaplib.IMAP4(MAIL_SERVER, nego_combo[1])
    elif nego_combo[0] == "starttls":
        context = ssl.create_default_context()
        imapclient = imaplib.IMAP4(MAIL_SERVER, nego_combo[1])
        imapclient.starttls(ssl_context=context)
    elif nego_combo[0] == "ssl":
        context = ssl.create_default_context()
        imapclient = imaplib.IMAP4_SSL(
            MAIL_SERVER, nego_combo[1], ssl_context=context)

    if imapclient is None:
        print('ERROR')
        exit()

    # 接続を開始する
    imapclient.login(MAIL_USER, MAIL_PASS)

    imapclient.select()
    typ, data = imapclient.search(None, "ALL")
    datas = data[0].split()

    fetch_num = MAX_FETCH_NUM
    if len(datas) < fetch_num:
        fetch_num = len(datas)

    print('len(data) = ', len(datas))
    print('MAX_FETCH_NUM = ', MAX_FETCH_NUM)
    print('fetch_num = ', fetch_num)

    msg_list = []
    for num in datas[len(datas) - fetch_num:]:
        typ, data = imapclient.fetch(num, '(RFC822)')
        msg = email.message_from_bytes(data[0][1])
        msg_list.append(msg)

    # 接続を閉じる
    imapclient.close()
    imapclient.logout()

    print('len(msg_list) = ', len(msg_list))

    return msg_list


if __name__ == '__main__':
    msg_list = receive_mail()

    for msg in msg_list:
        from_addr = str(make_header(decode_header(msg["From"])))
        subject = str(make_header(decode_header(msg["Subject"])))

        if msg.is_multipart() is False:
            body = msg.get_payload(decode=True)
            charset = msg.get_content_charset()
            if charset is not None:
                body = body.decode(charset, "ignore")

        else:
            body = ''
            for part in msg.walk():
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                charset = part.get_content_charset()
                if charset is not None:
                    payload = payload.decode(charset, "ignore")
                body += str(payload)

        print(subject + " <" + from_addr + ">")
