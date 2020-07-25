import time
import platform
import os
import glob
import webbrowser

import chardet
import pandas as pd
import extract_msg

TIME_INTERVAL = 0.2
USE_CATEGORIES = False
VERIFY_EMAIL_PATH = "verification"
INTRO_HTML = ('<html><body><div style = "text-align: center; position: fixed; left:20%; right:20%">\n'
              '<div style = "text-align: center; font-family: Calibri, Helvetica, sans-serif; color: rgb(0, 0, 0);">\n'
              '<h2>You can preview the emails here. '
              'Once done, close this window and go back to the terminal/command prompt.</h2>\n'
              '<h3><a href="email_1.html">Preview emails</a></h3>\n'
              '</div></div>\n'
              '</body></html>')
PREV_HTML = '<a href="<<prev_link>>">Prev</a>'
NEXT_HTML = '<a href="<<next_link>>">Next</a>'
HEADER_HTML = ('<div style = "width:60%; text-align:left; padding-left:20%; padding-right:20%">\n'
               '<html><body>\n'
               '<table style="width:100%"><tr>\n'
               '<td style="width:50%; text-align:left"><<prev>></td>\n'
               '<td style="width:50%; text-align:right"><<next>></td>\n'
               '</tr></table>\n'
               '<div style="font-family: Calibri, Helvetica, sans-serif; color: rgb(0, 0, 0);">\n'
               '<h3>Email no. <<email_num>> out of <<total_num>><br>\n'
               'Category: <<category>></h3>\n'
               '<h2>Subject: <<subject>></h2>\n'
               '<h3>To: <<recipient_names>><br>\n'
               'To Emails: <<recipient_emails>><br>\n'
               'CC: <<cc_names>><br>\n'
               'CC Emails: <<cc_emails>></h3>\n'
               '<hr>\n'
               '</div>\n'
               '</body></html>')
FOOTER_HTML = "</div>"
INTRO_TEXT = ("\nWelcome to Python Mail Merge for Outlook! v0.6 Published 29 June 2020\n"
              "I require an Excel spreadsheet containing the data and a '.msg' file(s)"
              "that contain the message template(s). "
              "I assume that the Excel spreadsheet contains a 'Name' and 'Email' column.")


def main():
    print(INTRO_TEXT)
    df, excel_filename = input_excel_file()
    columns, use_categories = process_excel_file(df)
    message_templates = get_message_templates(df, use_categories)
    emails_to_send = merge_emails(columns, df, message_templates, use_categories)
    verify_emails(emails_to_send)
    can_send = confirm_whether_to_send()
    if can_send:
        send_emails(emails_to_send, df)
        df.to_excel(excel_filename, index=False)


def input_excel_file():
    excel_filename = input("First, please tell me the file name of the Excel spreadsheet (it shouldn't be open): ")
    while True:
        try:
            df = pd.read_excel(excel_filename)
            break
        except FileNotFoundError:
            print("OOPS: That file doesn't exist.")
            excel_filename = input("Please try again: ")
    while True:
        try:
            with open(excel_filename, 'a'):
                pass
                break
        except PermissionError:
            raise RuntimeError("Please close the Excel spreadsheet file and try again.")
    return df, excel_filename


def process_excel_file(df):
    df.drop(df.columns[df.columns.str.contains('unnamed', case=False)], axis=1, inplace=True)
    columns = list(df.columns)
    if "Mail Merge Status" in columns:
        for value in df["Mail Merge Status"].unique():
            if value not in ["Sent", "ERROR", "Not sent yet", "Not sure"]:
                raise ValueError("Mail Merge Status column is not correct.")
        print("There is a 'Mail Merge Status' column in the spreadsheet. "
              "Therefore, I will only be sending emails for those that are not marked as 'sent'.")
    else:
        df.insert(0, "Mail Merge Status", "Not sent yet")
    if "Mail Merge Category" in columns:
        use_categories = True
    else:
        use_categories = False
    return columns, use_categories


def get_message_templates(df, use_categories):
    message_templates = {}
    if use_categories:
        categories = sorted(df["Mail Merge Category"].unique())
        print("\nI notice that there are the following categories in the spreadsheet:")
        for category in categories:
            print(f"\t- {category}")
        print(
            "I will assume that you will be sending different email templates based on the categories. "
            "You need to provide multiple '.msg' files")
        for category in categories:
            msg_file = input(f"Please tell me the '.msg' file for category '{category}': ")
            while True:
                try:
                    message_templates[category] = extract_text_from_msg(msg_file)
                    break
                except FileNotFoundError:
                    print("OOPS: That file doesn't exist.")
                    msg_file = input(f"Please try again for the category '{category}': ")
    else:
        msg_file = input(f"Please tell me the file name for the '.msg' file: ")
        while True:
            try:
                message_templates["No category"] = extract_text_from_msg(msg_file)
                break
            except FileNotFoundError:
                print("OOPS: That file doesn't exist.")
                msg_file = input(f"Please try again: ")
    return message_templates


def extract_text_from_msg(path):
    msg = extract_msg.Message(path)
    subject = msg.subject
    plain_text = msg.body
    html_bytes = msg.htmlBody
    encoding = chardet.detect(html_bytes)["encoding"]
    html_text = html_bytes.decode(encoding)
    return subject, plain_text, html_text  # return the subject, plain text message, and html message


def merge_emails(columns, df, message_templates, use_categories):
    emails_to_send = {}

    # TODO: Don't assume the name of the column and ask for user input??

    if "Email" in df.columns and "Emails" in df.columns:
        raise RuntimeError("There shouldn't be both an 'Email' column and 'Emails' column")
    elif "Emails" in df.columns:
        email_column = "Emails"
    elif 'Email' in df.columns:
        email_column = "Email"
    else:
        raise RuntimeError("There should be an 'Email' or 'Emails' column")

    if "Name" in df.columns and "Names" in df.columns:
        raise RuntimeError("There shouldn't be both a 'Name' column and 'Names' column")
    elif "Names" in df.columns:
        name_column = "Names"
    elif 'Name' in df.columns:
        name_column = "Name"
    else:
        raise RuntimeError("There should be a 'Name' or 'Names' column")

    cc_column = "CC Email"
    cc_names_column = "CC Names"

    num_rows = len(df)
    for i in range(num_rows):
        if df["Mail Merge Status"][i] != "Sent":
            if use_categories:
                category = df["Mail Merge Category"][i]
            else:
                category = "No category"
            subject, plain_text, html_text = message_templates[category]

            recipient_emails = [email.strip() for email in df[email_column][i].split(";")]
            recipient_names = [name.strip() for name in df[name_column][i].split(";")]
            if cc_column in df.columns:
                cc_emails = [email.strip() for email in df[cc_column][i].split(";")]

            for column in columns:
                replacement = str(df[column][i])
                if replacement == "nan":
                    replacement = ""
                replacement = replacement.replace("\n", "<br>")
                html_text = html_text.replace(f"<<{column}>>", replacement)
                html_text = html_text.replace(f"&lt;&lt;{column}&gt;&gt;", replacement)
                plain_text = plain_text.replace(f"<<{column}>>", replacement)
                plain_text = plain_text.replace(f"&lt;&lt;{column}&gt;&gt;", replacement)
                subject = subject.replace(f"<<{column}>>", replacement)
                subject = subject.replace(f"&lt;&lt;{column}&gt;&gt;", replacement)
            emails_to_send[i] = [category, subject, recipient_names, recipient_emails, cc_emails, plain_text, html_text]
    return emails_to_send


def verify_emails(emails_to_send):
    total_num = len(emails_to_send)
    try:
        if os.path.isdir(VERIFY_EMAIL_PATH):
            html_files = glob.glob(os.path.join(VERIFY_EMAIL_PATH, "email_*.html"))
            for file in html_files:
                os.remove(file)
        else:
            os.makedirs(VERIFY_EMAIL_PATH)
    except PermissionError:
        raise RuntimeError("I need to delete the verification folder. Please close the open email files.")
    for n, (_, args) in enumerate(emails_to_send.items()):
        verify_email(total_num, n + 1, args)
    path = os.path.join(VERIFY_EMAIL_PATH, "email_0.html")
    with open(path, 'w') as file:
        file.write(INTRO_HTML)
    print("I am now opening the preview file for you. "
          "If it doesn't open up, please navigate to the 'verification' folder and open 'email_0.html' in a browser.")
    if platform.system() == "Darwin":
        webbrowser.get("safari").open("file:///" + os.path.join(os.getcwd(), VERIFY_EMAIL_PATH, "email_0.html"), new=1)
    else:
        webbrowser.open(os.path.join(VERIFY_EMAIL_PATH, "email_0.html"), new=1)


def verify_email(total_num, n, args):
    prev_num = n - 1
    next_num = n + 1
    html = get_verify_email_html(n, total_num, prev_num, next_num, *args)
    filename = f"email_{n}.html"
    html_file_path = os.path.join(VERIFY_EMAIL_PATH, filename)
    with open(html_file_path, 'w') as html_file:
        html_file.write(html)


def get_verify_email_html(email_num, total_num, prev_num, next_num, category, subject, recipient_names,
                          recipient_emails, cc_emails, plain_text, html_text):
    header = HEADER_HTML
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
    header = header.replace("<<recipient_names>>", "; ".join(recipient_names))
    header = header.replace("<<recipient_emails>>", "; ".join(recipient_emails))
    header = header.replace("<<cc_names>>", "")
    header = header.replace("<<cc_emails>>", ";".join(cc_emails))
    html = header + html_text + FOOTER_HTML
    return html


def get_sent_status_mac(message):
    return message.was_sent.get()


def get_sent_status_windows(message):
    return message.sent


def confirm_whether_to_send():
    print("\nShould I send the emails?")
    continue_input = input("Type 'y' to send and 'n' to stop: ")
    while continue_input not in ["y", "n", "Y", "N"]:
        continue_input = input("Oops, please type 'y' or 'n': ")
    if continue_input in ["n", "N"]:
        return False
    else:
        return True


def send_emails(emails_to_send, df):
    print("\nOkay I will start sending emails now.\n")
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
                print(
                    f"ERROR: Email {i + 1}: '{subject}' to {recipient_names} took too long to respond. Likely not sent?")  # TRY DELETING THE EMAIL IN OUTLOOK?
                df.loc[i, "Mail Merge Status"] = "Not sure"
        to_delete = []
        for i, message in messages.items():
            try:
                tmp = get_sent_status(message)
                # print(f"Email {str(i)} not yet sent")
            except:  # Future: Identify the correct exception for windows and mac (for the email object being gone)
                subject = emails_to_send[i][1]
                recipient_names = "; ".join(emails_to_send[i][2])
                # print(f"Email {i + 1}: '{subject}' to {recipient_names} has been sent.")
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
    # mail.Body = plain_text
    mail.HTMLBody = html_text
    mail.Send()
    return mail


if __name__ == "__main__":
    if platform.system() == "Windows":
        import win32com.client as win32
        send_email = send_email_windows
        get_sent_status = get_sent_status_windows
        outlook = win32.Dispatch('outlook.application')
    elif platform.system() == "Darwin":
        from appscript import app, k
        send_email = send_email_mac
        get_sent_status = get_sent_status_mac
        outlook = app('Microsoft Outlook')
    else:
        raise SystemError("Cannot identify OS. I require windows or mac.")
    main()
