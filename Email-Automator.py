"""
Author: Mike Files

IMPORTANT:
- Ensure that Outlook Classic is installed and opened prior to running
- iCloud uses TLS, Gmail uses SSL and Outlook uses Outlook
- You're required to allow 'Programmatic Access' otherwise Outlook will throw a prompt on each attempt:
-- File -> Options -> Trust Center
-- Trust Center Settings
-- Click "Never warn me..." [You may need to Run as Admin]
"""

# ! SECTION: Read Emails into List
# Section-applicable imports
import csv

# Open CSV file
def load_email_list(filename):
	with open(filename, "r") as file:
		csv_reader = csv.reader(file)
		data = list(csv_reader)
	return data

# TODO: Change to Production email list when ready
email_list = load_email_list("emailer_utils\\PROD-Realtor_Email_List.csv")


# ! SECTION: Read Scripts into Dictionary
# Section-applicable imports
import json

# Open JSON file
def load_scripts(filename):
	with open(filename, "r") as file:
		data = json.load(file)
	return data

email_script = load_scripts("emailer_utils\\PROD-Email_Scripts.json")

# ! SECTION: Build email subject/body (as applicable)
import re
import random

def email_body_builder(recipient, intro, main, scenario, segue, offer, cta, addtl_info, ps, ending):
	email_message = ""
	base_intro = random.choice(intro)
	specific_intro = re.sub(r"\$\bNAME\b", recipient, base_intro)
	email_message +=specific_intro
	email_message += random.choice(main)
	email_message += random.choice(scenario)
	email_message += random.choice(segue)
	email_message += random.choice(offer)
	email_message += random.choice(cta)
	email_message += random.choice(addtl_info)
	email_message += random.choice(ps)
	email_message += random.choice(ending)
	return email_message

def email_subj_builder(recipient, subject):
	email_subject = ""
	base_subject = random.choice(subject)
	specific_subject = re.sub(r"\$\bNAME\b", recipient, base_subject)
	email_subject += specific_subject
	return email_subject

# ! SECTION: Email via Python
# Section-applicable imports
from emailer_utils import email_functions as ef
from emailer_utils import credentials as acct
import time

# ! SECTION: Loop through and send emails
# Declare/define a counter at 0
email_counter = 0

# Loop through slice
# TODO: Update the slice, by the thousand, each day
for contact in email_list[:1000]:
	# Set the contact information
	contact_name = contact[0]
	contact_email = contact[1]

	# Use modulo to determine the email account to use (five total options)
	email_acct_selector = email_counter % 5

	# Set email criteria (Outlook evaluates to "" because it does not use SSL/TLS)
	email_acct_sender = acct.credentials[email_acct_selector]["sender_email"]
	email_addr_used = acct.credentials[email_acct_selector]["username"]
	email_acct_password = acct.credentials[email_acct_selector]["app_password"]
	email_acct_server = acct.credentials[email_acct_selector]["smtp_server"]
	email_acct_port = acct.credentials[email_acct_selector]["port"]
	
	# Status update
	print("STATUS: Pending - FROM: {} - TO: {} - COUNT: {}".format(email_addr_used, contact_email, email_counter))

	if (email_addr_used == "mikefiles@me.com"):
		# Set the customized email message
		email_message = email_body_builder(
			contact_name,
			email_script["mike"]["intro"], email_script["mike"]["main"], email_script["mike"]["scenario"],
			email_script["mike"]["segue"], email_script["mike"]["offer"], email_script["mike"]["cta"],
			email_script["mike"]["additional_info"], email_script["mike"]["ps"], email_script["mike"]["ending"]
			)
		
		# Set the customized email subject
		email_subject = email_subj_builder(contact_name, email_script["mike"]["subject"])
		
		# Send email via TLS
		ef.tls_emailer(
			email_message, email_subject, email_acct_sender,
			contact_email, email_addr_used, email_acct_password,
			email_acct_server, email_acct_port
			)
	
	elif (email_addr_used == "alyssa.files@icloud.com"):
		# Set the customized email message
		email_message = email_body_builder(
			contact_name,
			email_script["alyssa"]["intro"], email_script["alyssa"]["main"], email_script["alyssa"]["scenario"],
			email_script["alyssa"]["segue"], email_script["alyssa"]["offer"], email_script["alyssa"]["cta"],
			email_script["alyssa"]["additional_info"], email_script["alyssa"]["ps"], email_script["alyssa"]["ending"]
			)
		
		# Set the customized email subject
		email_subject = email_subj_builder(contact_name, email_script["alyssa"]["subject"])
		
		# Send email via TLS
		ef.tls_emailer(
			email_message, email_subject, email_acct_sender,
			contact_email, email_addr_used, email_acct_password,
			email_acct_server, email_acct_port
			)
	
	elif (email_addr_used == "m.kefiles@gmail.com"):
		# Set the customized email message
		email_message = email_body_builder(
			contact_name,
			email_script["mike"]["intro"], email_script["mike"]["main"], email_script["mike"]["scenario"],
			email_script["mike"]["segue"], email_script["mike"]["offer"], email_script["mike"]["cta"],
			email_script["mike"]["additional_info"], email_script["mike"]["ps"], email_script["mike"]["ending"]
			)
		
		# Set the customized email subject
		email_subject = email_subj_builder(contact_name, email_script["mike"]["subject"])
		
		# Send email via SSL
		ef.ssl_emailer(
			email_message, email_subject, email_acct_sender,
			contact_email, email_addr_used, email_acct_password,
			email_acct_server, email_acct_port
		)
	
	elif (email_addr_used == "alyssa.m.files@gmail.com"):
		# Set the customized email message
		email_message = email_body_builder(
			contact_name,
			email_script["alyssa"]["intro"], email_script["alyssa"]["main"], email_script["alyssa"]["scenario"],
			email_script["alyssa"]["segue"], email_script["alyssa"]["offer"], email_script["alyssa"]["cta"],
			email_script["alyssa"]["additional_info"], email_script["alyssa"]["ps"], email_script["alyssa"]["ending"]
			)
		
		# Set the customized email subject
		email_subject = email_subj_builder(contact_name, email_script["alyssa"]["subject"])
		
		# Send email via SSL
		ef.ssl_emailer(
			email_message, email_subject, email_acct_sender,
			contact_email, email_addr_used, email_acct_password,
			email_acct_server, email_acct_port
		)
	
	elif (email_addr_used == "michael.files@outlook.com"):
		# Set the customized email message
		email_message = email_body_builder(
			contact_name,
			email_script["mike"]["intro"], email_script["mike"]["main"], email_script["mike"]["scenario"],
			email_script["mike"]["segue"], email_script["mike"]["offer"], email_script["mike"]["cta"],
			email_script["mike"]["additional_info"], email_script["mike"]["ps"], email_script["mike"]["ending"]
			)
		
		# Set the customized email subject
		email_subject = email_subj_builder(contact_name, email_script["mike"]["subject"])
		
		# Send email via Outlook
		ef.outlook_emailer(
			email_message, email_subject, contact_email
		)
	
	else:
		print("Error encountered. Email address called is not found")
		break

	# Increment email counter
	email_counter += 1

	# Brief pause between emails
	time.sleep(8)

print("~~~~~~~~~~~~~~~~~~~")
print("\nTODO - Mike to update last email sent in script!\n")
print("~~~~~~~~~~~~~~~~~~~")
# ! IMPORTANT - Last Send Date: mm/dd/yyyy
# ! IMPORTANT - Last Emailed: X through Y