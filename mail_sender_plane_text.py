import smtplib
import email
import msc

info_message_events = []
info_message_events.append('закрыты все процессы')
info_message_events.append('свободно места на диске')
info_message_events.append('eMail was sended by Python 3')

# список соединённый в текст
msg_body = '\r\n'.join(info_message_events)

# создание объекта "сообщение"
msg = email.message.EmailMessage()

# создание заголовков
msg.set_content('some text')
msg.set_type('text/plain; charset=utf-8')  # msg.set_type('text/plain;') msg.set_charset('utf-8')
msg['Date'] = email.utils.formatdate(localtime=True)
msg['Subject'] = msc.msc_msg_subject
msg['From'] = msc.msc_from_address
msg['To'] = msc.msc_to_address

msg.set_payload(msg_body.encode())

# печать всего пакета сообщения
print()
print('___msg___')
print(msg)

# отправка письма
smtp_link = smtplib.SMTP_SSL(msc.msc_mail_server)
smtp_link.login(msc.msc_login_user, msc.msc_login_pass)
smtp_link.send_message(msg, msc.msc_from_address, msc.msc_to_address)
smtp_link.quit()

print()
print('eMail was sended')
print()
print(f'закрыты процессы winrar, свободно места на диске с архивами = 50 G ... и статистика на почту отправлена')
