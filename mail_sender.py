import smtplib, email.message, email.utils, msc

info_message_events = []
info_message_events.append('закрыты все процессы')
info_message_events.append('свободно места на диске')
info_message_events.append('eMail was sended by Python 3')
print()
print('___info_message_events___')
print(info_message_events)

msg_post = '\r\n'.join(info_message_events)
print()
print('___msg_post___')
print(msg_post)

msg = email.message.EmailMessage()
print()
print('___msg___')
print(msg)

msg['Date'] = email.utils.formatdate(localtime=True)
msg['Subject'] = msc.msc_msg_subject
msg['From'] = msc.msc_from_address
msg['To'] = msc.msc_to_address
msg.set_content(msg_post)
msg['Body'] = 'msg_post msg_post msg_post msg_post msg_post'
msg.set_type('text/plain; charset=utf-8')

print()
print('___msg___')
print(msg)
print()
print('___dir(msg)___')

print()
print('___msg.values()___')
print(msg.values())
print()
print('___msg.get_params()___')
print(msg.get_params())
exit()

smtp_link = smtplib.SMTP_SSL(msc.msc_mail_server)
smtp_link.login(msc.msc_login_user, msc.msc_login_pass)
smtp_link.sendmail(msc.msc_from_address, msc.msc_to_address, msc.msc_msg)
smtp_link.quit()

print()
print('eMail was sended')

print()
print(f'закрыты процессы winrar, свободно места на диске с архивами = 50 G ... и статистика на почту отправлена')

