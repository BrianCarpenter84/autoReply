import imaplib
import email
from email.message import EmailMessage
from email.utils import make_msgid
import smtplib
import html
import time
from datetime import datetime


password = input("Enter password: ")
print(" \n" * 80)

monitor_inbox = True
while monitor_inbox:

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print("**********************************************************************")
    print(now)

# ==== Get emails ===============================================================

    with open("TextDocs/i.txt") as file:
        imap_server = file.read()

    with open("TextDocs/e.txt") as file:
        email_address = file.read()


    print(f"...Getting emails from account: {email_address}...\n")
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(email_address, password)
    imap.select("inbox")
    _, msgnums = imap.search(None, "ALL")

#  ========== Open contacted emails and save to a list ========================

    CONTACTED_EMAILS = []
    with open("TextDocs/CONTACTED_EMAILS.txt","r") as file:
        content = file.read()
        split_content = content.split("\n")
        for i in split_content:
            if i == " " or i == "":
                pass
            else:
                CONTACTED_EMAILS.append(i)

# =============== Get all emails from inbox and format them / save all variables ==============

    with open("TextDocs/s.txt") as file:
        site_submission = file.read()

    for msgnum in msgnums[0].split():
        _, data = imap.fetch(msgnum, "(RFC822)")
        message = email.message_from_bytes(data[0][1])
        if message.get("From") == site_submission:
            msg_list = []
            for part in message.walk():
                if part.get_content_type() == "text/html":
                    html_body = part.get_payload(decode=True).decode()
                    msg_body = html.unescape(html_body).replace('\r', '').replace('\n', ' ')
                    for i in msg_body:
                        msg_list = msg_body.split(' ')

                first_name = msg_list[1601]
                last_name = msg_list[1606]
                user_email = msg_list[1608]
                zip_code = msg_list[1613]

                if msg_list[1622] == "Yes":
                    send_jobs_link = True
                else:
                    send_jobs_link = False

#================= Format user message =====================================================

                msg_starting_index = msg_list.index('<br><strong>Message</strong>:')
                msg_ending_index = msg_list.index('<br></td></tr><tr><td')
                final_msg_list = []
                msg_item = msg_starting_index + 1
                while msg_item < msg_ending_index:
                    final_msg_list.append(msg_list[msg_item])
                    msg_item += 1
                formatted_msg = " ".join(final_msg_list)

#====== If jobs link requested, and user_email not in Contacted_emails, send reply and update archive===

                if user_email not in CONTACTED_EMAILS:
                    terminal_archive = f"""From: {message.get('From')}
To: {message.get('To')}
BCC: {message.get('BCC')}
Date: {message.get('Date')}
Subject: {message.get('Subject')}\n
-----------------------------------
Date/Time: {now}
sender name: {first_name} {last_name}
sender email: {user_email}
sender zip code: {zip_code}
send jobs link?: {send_jobs_link}
message:\n  {formatted_msg}"""

                    with open(f"TerminalArchive/{first_name}_{last_name}", "w") as file:
                        file.write(terminal_archive)

                    with open("TextDocs/referral_tracker.csv", "r") as file:
                        content = file.read()
                        csv_info = content.split("\n")
                    if not csv_info:
                        for i in csv_info:
                            if i[3] != user_email:
                                with open("TextDocs/referral_tracker.csv", "w") as file:
                                    date = message.get('Date').split(" ")
                                    date_list = []
                                    for i in date:
                                        date_list.append(i)
                                    date = f"{date_list[2]} {date_list[1]} {date_list[3]} {date_list[4]}"
                                    csv_info.append(f"{date},{first_name},{last_name},{user_email},{zip_code},{send_jobs_link}")
                                    file.write(f"\n{csv_info[csv_info.index(i)]}\n")
                    else:
                        with open("TextDocs/referral_tracker.csv", "w") as file:
                            date = message.get('Date').split(" ")
                            date_list = []
                            for i in date:
                                date_list.append(i)
                            date = f"{date_list[2]} {date_list[1]} {date_list[3]} {date_list[4]}"
                            csv_info.append(f"{date},{first_name},{last_name},{user_email},{zip_code},{send_jobs_link}")
                            for i in csv_info:
                                file.write(f"{csv_info[csv_info.index(i)]}\n")

                    attachment = "Logo/email_Logo_small.png"

                    if send_jobs_link:

                        html_body = f"""
                    <html>
                        <body>
                            <p><b>Hello {first_name}!</b></p> 
                            <p>Thank you for contacting me regarding your interest in working for <b>Penske Logistics</b>.</p> 
                            <p>This is an automated response, but I will reach out soon!</p> 
                            <p>But in the meantime, feel free to explore Penske Logistics job openings in your area with this link!</p> 
                            <a href="https://penske.jobs/jobs/?q=truck+driver&r=25&location={zip_code}">Penske.jobs</a> 
                            <p>The jobs listed will be within 25 miles of your zip code, but you can extend that distance if you like.</p> 
                            <p>If you see an opening that you meet the minimum requirements for, you can apply now!</p> 
                            <p>Just remember to include my name, <b>Brian Carpenter</b> , in the <b>Driver Referral</b> field!</p> 
                            <p>Thanks again, and I'm looking forward to talking with you.</p>

                            <p>Regards,</p> 
                            <p>Brian Carpenter</p> 
                            <p><b>CLETrucker</b></p> 

                        </body>
                    </html>
                    """
                    else:
                        html_body = f"""
                                    <html>
                                        <body>
                                            <p><b>Hello {first_name}!</b></p> 
                                            <p>Thank you for contacting me.</p> 
                                            <p>This is an automated response.</p> 
                                            <p>I just wanted to let you know that I have received your message and that I will reach out soon!</p> 
                                            <p>Thanks again, and I'm looking forward to talking with you.</p>

                                            <p>Regards,</p> 
                                            <p>Brian Carpenter</p> 
                                            <p><b>CLETrucker</b></p> 

                                        </body>
                                    </html>
                                    """
                    with open("TextDocs/from.txt") as file:
                        sent_from = file.read()

                    msg = EmailMessage()
                    msg["Subject"] = "Auto Reply CLETrucker"
                    msg["From"] = sent_from
                    msg["To"] = user_email

                    attachment_cid = make_msgid()

                    msg.set_content('<b>%s</b><br/><img src="cid:%s"/><br/>' % (html_body, attachment_cid[1:-1]), 'html')

                    with open(attachment, 'rb') as fp:
                        msg.add_related(
                            fp.read(), 'image', 'png', cid=attachment_cid)

                    with open("TextDocs/smtp.txt") as file:
                        host = file.read()

                    with smtplib.SMTP_SSL(host, 465) as server:
                        server.ehlo()
                        server.login(f"{email_address}", f"{password}")
                        server.send_message(msg,sent_from, f"{user_email}")
                        server.close()
                    CONTACTED_EMAILS.append(user_email)
                    print(f"{user_email} has been contacted. TerminalArchive has been updated.")
                else:
                    print(f"{user_email} has been previously contacted. No email sent.")
        else:
            pass
    imap.close()

    if msgnums == [b'']:
        print("No new messages.\n")

    with open("TextDocs/CONTACTED_EMAILS.txt", "w") as file:
        for i in CONTACTED_EMAILS:
            file.write(f"{i}\n")
    print("\n...Loop complete. Waiting to restart...\n**********************************************************************\n \n")
    time.sleep(60)


