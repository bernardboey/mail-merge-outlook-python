import time
import platform
import os
import email
import glob
import webbrowser
import chardet

import pandas as pd
import extract_msg

TIME_INTERVAL = 0.2
USE_CATEGORIES = False
VERIFY_EMAIL_PATH = "verification"

def extract_text_from_msg(path):
	msg = extract_msg.Message(path)
	subject = msg.subject
	plain_text = msg.body
	html_bytes = msg.htmlBody
	encoding = chardet.detect(html_bytes)["encoding"]
	html_text = html_bytes.decode(encoding)
	return subject, plain_text, html_text  # return the subject, plain text message, and html message


def extract_text_from_email(path):
	"""
	This function takes in a .eml file name (outlook email file) and returns
	the subject, and message text from the email
	"""
	with open(path, 'r') as file:  # read .eml file
		msg = email.message_from_file(file)  # parse .eml file into message object
		subject = msg["Subject"]  # store the subject of the email in a variable
		plain_text = None  # initialise the plain text message as NULL
		html_text = None  # initialise the html message as NULL
		for part in msg.walk():  # "walk" through the sub parts of the .eml file
			if(part.get_content_type() == "text/plain"):  # find the part contains the plain text message
				plain_text = part.get_payload()  # store the plain text message in variable
			if(part.get_content_type() == "text/html"):  # find the part contains the html message
				html_text = part.get_payload()  # store the html message in variable
		if plain_text is None and html_text is None:  # check if the plain text and html messages are missing
			raise ValueError("Cannot find plain text/html")  # if yes, stop the program
		plain_text = plain_text.replace("=\n","\n")
		html_text = html_text.replace("=\n","\n")
		return subject, plain_text, html_text  # return the subject, plain text message, and html message


def get_sent_status_mac(message):
	tmp = message.was_sent.get()


def get_sent_status_windows(message):
	tmp = message.sent


def send_emails(emails_to_send, df):
	num_messages = len(emails_to_send)
	num_sent = (df["Mail Merge Status"].values == 'Sent').sum()
	messages = {}
	print("Sending emails...")
	for i, args in emails_to_send.items():
		subject, recipient_names = args[1], "; ".join(args[2])
		try:
			print(f"Processing email {i + 1}: '{subject}' to {recipient_names}......")
			messages[i] = send_email(*args)
		except:
			df.loc[i, "Mail Merge Status"] = "ERROR"
			print(f"ERROR: Email {i + 1}: '{subject}' to {recipient_names} couldn't be sent.")
	total_time = 0
	while len(messages) != 0:
		time.sleep(TIME_INTERVAL)
		total_time += TIME_INTERVAL
		if total_time > 300:
			for i, message in messages.items():
				subject = emails_to_send[i][1]
				recipient_names = "; ".join(emails_to_send[i][2])
				print(f"ERROR: Email {i + 1}: '{subject}' to {recipient_names} took too long to respond. Likely not sent?") # TRY DELETING THE EMAIL IN OUTLOOK?
				df.loc[i, "Mail Merge Status"] = "Not sure"
		to_delete = []
		for i, message in messages.items():
			try:
				tmp = get_sent_status(message)
				#print(f"Email {str(i)} not yet sent")
			except:
				subject = emails_to_send[i][1]
				recipient_names = "; ".join(emails_to_send[i][2])
				#print(f"Email {i + 1}: '{subject}' to {recipient_names} has been sent.")
				df.loc[i, "Mail Merge Status"] = "Sent"
				to_delete.append(i)
		for i in to_delete:
			del messages[i]
	num_sent = (df["Mail Merge Status"].values == 'Sent').sum() - num_sent
	print(f"Finished sending. {num_sent} out of {num_messages} emails successfully sent.")


def send_email_mac(category, subject, recipient_names, recipient_emails, cc_emails, plain_text, html_text):
	msg = outlook.make(
    	new=k.outgoing_message,
    	with_properties={
        	k.subject: subject,
        	k.content: html_text})
	for email_address in recipient_emails:
		msg.make(
	    	new=k.to_recipient,
	    	with_properties={
	        	k.email_address: {
	            	k.address: email_address}})
	for email_address in cc_emails:
		msg.make(
	    	new=k.cc_recipient,
	    	with_properties={
	        	k.email_address: {
	            	k.address: email_address}})
	msg.send()
	return msg


def send_email_windows(category, subject, recipient_names, recipient_emails, cc_emails, plain_text, html_text):
	mail = outlook.CreateItem(0)
	mail.To = "; ".join(recipient_emails)
	mail.CC = "; ".join(cc_emails)
	mail.Subject = subject
	#mail.Body = plain_text
	mail.HTMLBody = html_text
	mail.Send()
	return mail

def get_verify_email_html(email_num, total_num, prev_num, next_num, category, subject, recipient_names, recipient_emails, cc_emails, plain_text, html_text):
	PREV_HTML = '<a href="<<prev_link>>">Prev</a>'
	NEXT_HTML = '<a href="<<next_link>>">Next</a>'
	HEADER = """<div style = "text-align: left; padding-left:20%; padding-right:20%">
				<html><body>
				<table style="width:100%"><tr>
				<td style="width:50%; text-align:left"><<prev>></td>
				<td style="width:50%; text-align:right"><<next>></td>
				</tr></table>
				<div style="font-family: Calibri, Helvetica, sans-serif; color: rgb(0, 0, 0);">
				<h3>Email no. <<email_num>> out of <<total_num>><br>
				Category: <<category>></h3>
				<h2>Subject: <<subject>></h2>
				<h3>To: <<recipient_names>><br>
				To Emails: <<recipient_emails>><br>
				CC: <<cc_names>><br>
				CC Emails: <<cc_emails>></h3>
				<hr>
				</div>
				</body></html>"""
	FOOTER = "</div>"
	header = HEADER
	if email_num == 1:
		header = header.replace("<<prev>>", "")
	else:
		prev_html_formatted = PREV_HTML.replace("<<prev_link>>", f"email_{prev_num}.html")
		header = header.replace("<<prev>>", prev_html_formatted)
	if email_num == total_num:
		header = header.replace("<<next>>", "")
	else:
		next_html_formatted = NEXT_HTML.replace("<<next_link>>", f"email_{next_num}.html")
		header = header.replace("<<next>>", next_html_formatted)
	header = header.replace("<<email_num>>", str(email_num))
	header = header.replace("<<total_num>>", str(total_num))
	header = header.replace("<<category>>", category)
	header = header.replace("<<subject>>", subject)
	header = header.replace("<<recipient_names>>", ";".join(recipient_names))
	header = header.replace("<<recipient_emails>>", ";".join(recipient_emails))
	header = header.replace("<<cc_names>>", "")
	header = header.replace("<<cc_emails>>", ";".join(cc_emails))
	html = header + html_text + FOOTER
	return html

def verify_emails(emails_to_send):
	total_num = len(emails_to_send)
	try:
		if os.path.isdir(VERIFY_EMAIL_PATH):
			html_files = glob.glob(os.path.join(VERIFY_EMAIL_PATH,"email_*.html"))
			for file in html_files:
				os.remove(file)
		else:
			os.makedirs(VERIFY_EMAIL_PATH)
	except:
		raise RuntimeError("I need to delete the verification folder. Please close the open email files.")
	def verify_email(n, i, args):
		prev_num = n - 1
		next_num = n + 1
		html = get_verify_email_html(n, total_num, prev_num, next_num, *args)
		filename = f"email_{n}.html"
		path = os.path.join(VERIFY_EMAIL_PATH,filename)
		with open(path, 'w') as file:
			file.write(html)
	for n, (i, args) in enumerate(emails_to_send.items()):
		verify_email(n + 1, i, args)
	path = os.path.join(VERIFY_EMAIL_PATH,"email_0.html")
	with open(path, 'w') as file:
		intro_html = """<html><body><div style = "text-align: center; position: fixed; left:20%; right:20%">
						<div style = "text-align: center; font-family: Calibri, Helvetica, sans-serif; color: rgb(0, 0, 0);">
						<h2>You can preview the emails here. Once done, close this window and go back to the terminal/command prompt.</h2>
						<h3><a href='email_1.html'>Preview emails</a></h3>
						</div></div>
						</body></html>"""
		file.write(intro_html)
	print("I am now opening the preview file for you. If it doesn't open up, please navigate to the 'verification' folder and open 'email_0.html' in a browser.")
	if platform.system() == "Darwin":
		webbrowser.get("safari").open("file:///" + os.path.join(os.getcwd(),VERIFY_EMAIL_PATH,"email_0.html"), new=1)
	else:
		webbrowser.open(os.path.join(VERIFY_EMAIL_PATH,"email_0.html"), new=1)


def main():
	print("\nWelcome to Python Mail Merge for Outlook! v1.0 Published 24 May 2020")
	print("I require an Excel spreadsheet containing the data and a '.msg' file(s) that contain the message template(s).")
	filename = input("First, please tell me the file name of the Excel spreadsheet (it shouldn't be open): ")
	while True:
		try:
			df = pd.read_excel(filename)
			break
		except:
			print("OOPS: Either that file doesn't exist or there is an error.")
			filename = input("Please try again: ")
	while True:
		try:
			with open(filename, 'a') as file:
				pass
				break
		except:
			raise RuntimeError("Please close the Excel spreadsheet file and try again.")
	columns = list(df.columns)
	df.drop(df.columns[df.columns.str.contains('unnamed', case=False)], axis=1, inplace=True)
	if "Mail Merge Status" in columns:
		for value in df["Mail Merge Status"].unique():
			if value not in ["Sent", "ERROR", "Not sent yet", "Not sure"]:
				raise ValueError("Mail Merge Status column is not correct.")
		print("There is a 'Mail Merge Status' column in the spreadsheet. Therefore, I will only be sending emails for those that are not marked as 'sent'.")
	else:
		df.insert(0,"Mail Merge Status","Not sent yet")
	num_rows = len(df)
	if "Mail Merge Category" in columns:
		USE_CATEGORIES = True
		categories = df["Mail Merge Category"].unique()
		message_templates = {}
		print("\nI notice that there are the following categories in the spreadsheet:")
		for category in categories:
			print(f"\t- {category}")
		print("I will assume that you will be sending different email templates based on the categories. You need to provide multiple '.msg' files")
		for category in categories:
			msg_file = input(f"Please tell me the '.msg' file for category '{category}': ")
			while True:
				try:
					message_templates[category] = extract_text_from_msg(msg_file)
					break
				except:
					print("OOPS: Either that file doesn't exist or there is an error.")
					msg_file = input(f"Please try again for the category '{category}': ")
	else:
		msg_file = input(f"Please tell me the file name for the '.msg' file: ")
		while True:
			try:
				message_template = extract_text_from_msg(msg_file)
				break
			except:
				print("OOPS: Either that file doesn't exist or there is an error.")
				filename = input(f"Please try again: ")
	emails_to_send = {}
	for i in range(num_rows):
		if df["Mail Merge Status"][i] != "Sent":
			if USE_CATEGORIES:
				category = df["Mail Merge Category"][i]
				subject, plain_text, html_text = message_templates[category]
			else:
				subject, plain_text, html_text = message_template
				category = None
			recipient_emails = [df["Email"][i]] # SHOULD WE NOT ASSUME HERE?
			recipient_names = [df["Name"][i]] # SHOULD WE NOT ASSUME HERE?
			cc_emails = []

			for column in columns:
				replacement = str(df[column][i])
				replacement = replacement.replace("\n","<br>")
				html_text = html_text.replace(f"<<{column}>>", replacement)
				html_text = html_text.replace(f"&lt;&lt;{column}&gt;&gt;", replacement)
				plain_text = plain_text.replace(f"<<{column}>>", replacement)
				plain_text = plain_text.replace(f"&lt;&lt;{column}&gt;&gt;", replacement)
			emails_to_send[i] = [category, subject, recipient_names, recipient_emails, cc_emails, plain_text, html_text]

	verify_emails(emails_to_send)
	print("\nShould I send the emails?")
	check = input("Type 'y' to send and 'n' to stop: ")
	while check not in ["y", "n", "Y", "N"]:
		check = input("Oops, please type 'y' or 'n': ")
	if check in ["n", "N"]:
		exit()
	print("\nOkay I will start sending emails now.\n")
	send_emails(emails_to_send, df)

	df.to_excel(filename, index=False)

if __name__ == "__main__":
	if platform.system() == "Windows":
		send_email = send_email_windows
		get_sent_status = get_sent_status_windows
		import win32com.client as win32
		outlook = win32.Dispatch('outlook.application')
	elif platform.system() == "Darwin":
		send_email = send_email_mac
		get_sent_status = get_sent_status_mac
		from appscript import app, k
		outlook = app('Microsoft Outlook')
	else:
		raise SystemError("Cannot identify OS. Seems like not windows and not mac.")
	main()