# Client SDK for Google Sheets
import gspread
# OAuth2 client library to authenticate against Google APIs
from oauth2client.service_account import ServiceAccountCredentials
# Python's standard SMTP client to send emails
import smtplib
# Helper class for formatting email MIME bodies
from email.mime.text import MIMEText
# Date/time utilities
from datetime import date

# Credentials for Google Sheets - details are stored in client_secret.json
scope = ["https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("client_secret.json",
    scope)
client = gspread.authorize(creds)

# Opens connection to SMTP server
gmail_server = smtplib.SMTP("smtp.gmail.com", 587) # Host and port may vary for your email provider
gmail_server.ehlo() # Establish connection
gmail_server.starttls() # Request secure channel
# Replace with your credentials
gmail_user = "enter-gmail"
gmail_pw = "enter-password"
gmail_server.login(gmail_user, gmail_pw)

# Definition of team members - recipients of the emails
team = [
    {"name" : "recipient-name0", "email" : "recipient-email0"},
    {"name" : "recipient-name1", "email" : "recipient-email1"},
]

# Get Google Sheet contents and store in list "list"
sheet = client.open("Enter Sheet").sheet1
list = sheet.get_all_values()

# The subsequent code will create a personalised email containg the tasks a team member is responsible for according to a Google Sheet containing four columns:
# Area, Task, Name, Due

# Iterate through team members and their tasks respectively
# A members tasks will temporarily be stored in list "member_tasks"
for member in team:
    name = member["name"]
    email = member["email"]
    member_tasks = []
    for row in list:
        lead = row[2]
        if lead != name:
            continue
        task = {"area" : row[0], "task" : row[1], "due" : row[3]} # Row 2 contains the name and is omitted
        member_tasks.append(task)
    
    # Definition of email text for iteration's member
    # The body will contain a list of member's tasks (area, task and due date) according to list "member_tasks"
    email_text = """\
Hi %s,

this is an automated, weekley email, which informs you of your tasks according to Task List. Currently, you are responsible for the following tasks:

%s

Under the following link you can view the entire Task List:
enter-link-to-google-sheet

Replies to this email will not be read. If you have any questions, please contact enter-name at enter-email.
""" % (name, '\n\n'.join(map(lambda d: "Area: %s\nTask: %s\nDue: %s"
        % (d["area"], d["task"], d["due"]), member_tasks))) # Joins the tasks into a multi-line string.
    
    # Build email for iteration's member
    msg = MIMEText(email_text, "plain")
    msg["Subject"] = "Your Tasks " + date.today().strftime('%d.%m.%Y')
    msg["From"] = "Task List"
    msg["To"] = email

    # Send email for iteration's member
    gmail_server.sendmail(gmail_user, [email], msg.as_string())

# Close connection with Gmail
gmail_server.close()
