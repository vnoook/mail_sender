import smtplib, email.message, email.utils, msc

info_message_events = []
info_message_events.append('закрыты все процессы')
info_message_events.append('свободно места на диске')
info_message_events.append('eMail was sended by Python 3')

# список из строк
print()
print('___info_message_events___')
print(info_message_events)

# список соединённый в текст
msg_body = '\r\n'.join(info_message_events)
print()
print('___msg_body___')
print(msg_body)

# создание объекта "сообщение"
msg = email.message.EmailMessage()
print()
print('___msg___')
print(msg)

# создание заголовков
# msg['Date'] = email.utils.formatdate(localtime=True)
msg['Subject'] = msc.msc_msg_subject
msg['From'] = msc.msc_from_address
msg['To'] = msc.msc_to_address
msg.set_payload(msg_body)

# msg.as_string(msg)
# msg.set_type('text/plain')  # 'text/plain; charset=utf-8'
# msg['Body'] = msg_body


# печать всего пакета сообщения
print()
print('___msg___')
print(msg)

# печать свойств объекта сообщения
print()
print('___dir(msg)___')
print(dir(msg))

# печать значений обекта
print()
print('___msg.values()___')
print(msg.values())

# печать параметров объекта
print()
print('___msg.get_params()___')
print(msg.get_params())

# отправка письма
smtp_link = smtplib.SMTP_SSL(msc.msc_mail_server)
smtp_link.login(msc.msc_login_user, msc.msc_login_pass)
# smtp_link.sendmail(msc.msc_from_address, msc.msc_to_address, msg)
smtp_link.send_message(msg, msc.msc_from_address, msc.msc_to_address)
smtp_link.quit()

print()
print('eMail was sended')
print()
print(f'закрыты процессы winrar, свободно места на диске с архивами = 50 G ... и статистика на почту отправлена')
