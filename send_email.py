import smtplib, ssl
import win32com.client as win32
from datetime import datetime
import os


def smtp_setup():
    port = 465  # For SSL
    smtp_server = "smtp-mail.outlook.com"
    sender_email = "amelia.swensen@tigris-fp.com"  # Enter your address
    receiver_email = "amelia.swensen@tigris-fp.com"  # Enter receiver address
    password = input("Type your password and press enter: ")
    message = """\
    Subject: Hi there

    This message is sent from Python."""

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message)


def send_outlook_email_hardcoded():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = 'PC 123 Scorecard W29'
    mail.To = "amelia.swensen@tigris-fp.com; ameliaswen@gmail.com"
    attachment = mail.Attachments.Add(r'C:\Users\amelia.swensen\OneDrive - TigrisFP\Documents\CS Accounting Project Rankings_Planning.xlsx')
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "barbie_img")
    mail.HTMLBody = r"""
    Hi Barbie:<br><br>
    <img src="cid:barbie_img"><br><br>
    """
    mail.Attachments.Add(r'C:\Users\amelia.swensen\OneDrive - TigrisFP\Documents\hi_barbie.png')
    mail.Send()

def send_outlook_email(file_name, mail_to, file_path, html):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = file_name[:-5]
    mail.To = mail_to
    # mail.CC = 'amelia.swensen@tigris-fp.com'
    mail.HTMLBody = rf"""
        {html}<br>
        Thanks,<br><br>
        <b>Brian Davis</b><br>
        Tigris Fulfillment Partners <br>
        (919) 561-4009 <br>
        <a href="mailto:Brian.Davis@Tigris-FP.com">Brian.Davis@Tigris-FP.com</a><br>
        """
    mail.Attachments.Add(file_path)
    mail.Send()
    print('sent')