import imaplib
import email
import pandas as pd
import yaml

# Set your Gmail username and password in the credentials.yml file

# open the file
with open('credentials.yml', 'r') as f:
    content = f.read()

# load the content
my_credentials = yaml.load(content, Loader=yaml.FullLoader)

# get the username and password
username = my_credentials['user']
password = my_credentials['password']

# Connect to Gmail
my_email = imaplib.IMAP4_SSL("imap.gmail.com")
my_email.login(username, password)

# Select the "All Mail" folder
my_email.select("Inbox")

# Get a list of all emails in the folder
status, messages = my_email.search(None, "ALL")
message_ids = messages[0].split()

num_emails_to_fetch = 50
message_ids = message_ids[:num_emails_to_fetch]

# Create a list to store the sender name and email address
sender_info = []

# Iterate through the list of emails
for message_id in message_ids:
    # Get the email
    status, msg_data = my_email.fetch(message_id, "(RFC822)")
    
    # Check if the fetch was successful
    if status == 'OK':
        msg = email.message_from_bytes(msg_data[0][1])  # Convert bytes to email message

        # Get the sender name and email address
        data_dict = {}
        # This is a tupla
        information = email.utils.parseaddr(msg["From"])
        
        sender, email_address = information
        data_dict["sender"] = sender
        data_dict["Email_address"] = email_address

        # Add the sender information to the list
        sender_info.append(data_dict)

# Create a Pandas DataFrame from the list
df = pd.DataFrame(sender_info)

# Save the DataFrame to Excel - csv
df.to_csv("sender_info.csv", index=False)

# Logout from the email
my_email.logout()