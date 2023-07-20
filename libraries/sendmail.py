import smtplib  
import email.utils
from robot.api import logger
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
from variables import TODAY
from pathlib import Path
import os

def send(s_sender,s_sendername,s_recipient,s_cc,s_username,s_password,s_host,s_subject,s_bodytext,s_bodytexthtml, attachment_file_name):
    # Replace sender@example.com with your "From" address. 
    # This address must be verified.
        
    SENDER = s_sender  
    SENDERNAME = s_sendername
    #rem to test
    #s_recipient = "Johnson.Salim@contracted.sampoerna.com"
    
    # Replace recipient@example.com with a "To" address. If your account 
    # is still in the sandbox, this address must be verified.
    RECIPIENT  = s_recipient
    
    # CC
    CC = s_cc

    # Replace smtp_username with your Amazon SES SMTP user name.
    USERNAME_SMTP = s_username
    
    # Replace smtp_password with your Amazon SES SMTP password.
    PASSWORD_SMTP = s_password
    
    # (Optional) the name of a configuration set to use for this message.
    # If you comment out this line, you also need to remove or comment out
    # the "X-SES-CONFIGURATION-SET:" header below.
    # CONFIGURATION_SET = "ConfigSet"
    
    # If you're using Amazon SES in an AWS Region other than US West (Oregon), 
    # replace email-smtp.us-west-2.amazonaws.com with the Amazon SES SMTP  
    # endpoint in the appropriate region.
    HOST = s_host
    PORT = 587
    
    # The subject line of the email.
    SUBJECT = s_subject
    
    # The email body for recipients with non-HTML email clients.
    BODY_TEXT = s_bodytext
    # ("Amazon SES Test\r\n"
    #              "This email was sent through the Amazon SES SMTP "
    #              "Interface using the Python smtplib package."
    #             )
    
    # The HTML body of the email.
    BODY_HTML = s_bodytexthtml
    
    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = SUBJECT
    msg['From'] = email.utils.formataddr((SENDERNAME, SENDER))
    msg['To'] = RECIPIENT
    msg['CC'] = CC
    # Comment or delete the next line if you are not using a configuration set
    # msg.add_header('X-SES-CONFIGURATION-SET',CONFIGURATION_SET)
    
    # Record the MIME types of both parts - text/plain and text/html.
    # part1 = MIMEText(BODY_TEXT, 'plain')
    part2 = MIMEText(BODY_HTML, 'html')
    
    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    # msg.attach(part1)
    msg.attach(part2)
    
    #add attachment
    if attachment_file_name != None:
        part = MIMEBase('application', "octet-stream")
        with open(attachment_file_name, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="'+os.path.basename(attachment_file_name)+'"')
        msg.attach(part)


    
    # Try to send the message.
    try:  
        server = smtplib.SMTP(HOST, PORT)
        server.ehlo()
        server.starttls()
        #stmplib docs recommend calling ehlo() before & after starttls()
        server.ehlo()
        server.login(USERNAME_SMTP, PASSWORD_SMTP)
        server.sendmail(SENDER, RECIPIENT.split(',') + CC.split(','), msg.as_string())
        server.close()
    # Display an error message if something goes wrong.
    except Exception as e:
        # print("Error: ", e)
        logger.info("PY Log: " + str(e) + f" - on: {TODAY}")
        return 1
    else:
        # print ("Email sent!")
        return 0
