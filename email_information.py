import imaplib
import email
import pandas as pd
import yaml
import re

pattern = r'=\?UTF-8\?Q\?(.*?)\?='

# Set your Gmail username and password in the credentials.yml file

def get_credentials_from_yaml():
    with open('credentials.yml', 'r') as f:
        content = f.read()

    my_credentials = yaml.load(content, Loader=yaml.FullLoader)
    return my_credentials['user'], my_credentials['password']

def connect_to_gmail():
    username, password = get_credentials_from_yaml()
    my_email = imaplib.IMAP4_SSL("imap.gmail.com")
    my_email.login(username, password)
    return my_email

def fetch_sender_info(email_ids, my_email):
    sender_info = []
    for message_id in email_ids:
        status, msg_data = my_email.fetch(message_id, "(RFC822)")
        if status == 'OK':
            msg = email.message_from_bytes(msg_data[0][1])
            data_dict = {}
            information = email.utils.parseaddr(msg["From"])
            sender_name = extract_sender_name(information)
            data_dict["sender"] = sender_name[0]
            data_dict["Email_address"] = sender_name[1]
            sender_info.append(data_dict)
    return sender_info

def extract_sender_name(information):
    
    if information[0] == '':
        email = information[1]
        sender = information[1].split("@")[0] 
        return sender, email 
    
    match = re.search(pattern, information[0])
    if match:
        email = information[1]
        sender = match.group(1).split("=")[0].replace('_', ' ')
        return sender, email 

    return information

def save_sender_info_to_csv(sender_info):
    df = pd.DataFrame(sender_info)
    df.to_csv("sender_info.csv", index=False)

def main():
    # Connect to Gmail
    my_email = connect_to_gmail()

    # Select the "All Mail" folder
    my_email.select("Inbox")

    # Get a list of all emails in the folder
    status, messages = my_email.search(None, "ALL")
    message_ids = messages[0].split()

    num_emails_to_fetch = 30
    message_ids = message_ids[:num_emails_to_fetch]

    # Fetch sender information
    sender_info = fetch_sender_info(message_ids, my_email)

    # Save the sender information to CSV
    save_sender_info_to_csv(sender_info)

    # Logout from the email
    my_email.logout()

if __name__ == "__main__":
    main()
