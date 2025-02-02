import smtplib
import ssl
from email.mime.text import MIMEText
import sys
import win32com.client as win32

def tls_emailer(body, subj, sender, recipient, uname, pword, outbound_server, port):
	# Create a secure SSL context
	context = ssl.create_default_context()

	# Set credentials
	username = uname
	password = pword

	# Set outbound criteria
	smtp_server = outbound_server
	smtp_port = port

	# Get message data
	msg_body = body
	msg_from = sender
	msg_to = recipient
	msg_subject = subj

	# Set message data
	email_message = MIMEText(msg_body)
	email_message["From"] = msg_from
	email_message["To"] = msg_to
	email_message["Subject"] = msg_subject

	try:
		# Set server session
		server_session = smtplib.SMTP(smtp_server, smtp_port)
		server_session.starttls(context = context)
		server_session.login(username, password)

		# Send the email
		server_session.sendmail(
			email_message["From"],
			email_message["To"],
			email_message.as_string()
			)
		print("Message sent...")
	except Exception as e:
		print("Error sending email...")
		print("~~~~~~~~~~~~~~~~~~~~~~~")
		print(e)
		print("~~~~~~~~~~~~~~~~~~~~~~~")
		sys.exit()
	finally:
		server_session.quit()

def ssl_emailer(body, subj, sender, recipient, uname, pword, outbound_server, port):
	# Create a secure SSL context
	context = ssl.create_default_context()

	# Set credentials
	username = uname
	password = pword

	# Set outbound criteria
	smtp_server = outbound_server
	smtp_port = port

	# Get message data
	msg_body = body
	msg_from = sender
	msg_to = recipient
	msg_subject = subj

	# Set message data
	email_message = "Subject: " + msg_subject + "\n\n" + msg_body

	with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
		server.login(username, password)
		server.sendmail(
			from_addr=msg_from,
			to_addrs=msg_to,
			msg=email_message
			)
		print("Message sent...")

def outlook_emailer(body, subj, recipient):
	outlookApp = win32.Dispatch("Outlook.Application")
	mail = outlookApp.CreateItem(0)
	mail.To = recipient
	mail.Subject = subj
	mail.Body = body
	mail.Send()
