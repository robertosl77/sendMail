from sendMail import OutlookEmailSender

email_sender = OutlookEmailSender('sr.macros@gmail.com', '123465')
email_sender.send_custom_email(
    'sr.macros@gmail.com',
    'rsleiva@edenor.com',
    None,None,'Prueba','Prueba Body')
