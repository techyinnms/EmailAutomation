from flask import Flask, request, render_template
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

# Initialize Flask app
app = Flask(__name__)

# Configure the Google Sheets API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('email-auto-394611-52c51154893a.json', scope)
client = gspread.authorize(credentials)

# Your Google Spreadsheet details
SPREADSHEET_NAME = 'Gmail - Test'
WORKSHEET_NAME = 'Your Worksheet Name'

# Function to get data from the Google Spreadsheet
def get_spreadsheet_data():
    sheet = client.open(SPREADSHEET_NAME).get_worksheet(0)
    return sheet.get_all_records()

# Function to update the status in the Google Spreadsheet
def update_status(row_index, status):
    sheet = client.open(SPREADSHEET_NAME).get_worksheet(0)
    cell = sheet.cell(row_index, 3)  # Assuming status column is the 3rd column (C)
    cell.value = status
    sheet.update_cell(row_index, 3, status)

# Function to send the task update email
def send_task_update_email(sender_email, recipient_email, task_name, row_index):
    smtp_server = "smtp.gmail.com"  # Use your email provider's SMTP server if not using Gmail
    smtp_port = 25  # Use the appropriate SMTP port for your email provider

    # Replace the following with your own email credentials
    sender_password = "frqxgqdnqttnolfy"  # Replace with the sender email password

    msg = MIMEMultipart("alternative")
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = f"Task Update Request:"

    # Read task_update.html template
    with open("task_update.html") as f:
        template = f.read()

    # Append the row_index as a query parameter in the feedback link
    feedback_link = f"http://localhost:5000/feedback/{row_index}"
    template = template.format(task_name=task_name, row_index=row_index, feedback_link=feedback_link)

    msg.attach(MIMEText(template, "html"))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print(f"Email sent successfully to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}. Error: {e}")

# Route to render the homepage with sender email input form
@app.route('/')
def homepage():
    return render_template('index.html')

# Route to handle form submission and send the task update email
@app.route('/send_email', methods=['POST'])
def send_email():
    sender_email = request.form['sender_email']

    # Assuming your tasks are fetched from the Google Sheets
    tasks_data = get_spreadsheet_data()

    for row_index, task in enumerate(tasks_data, start=2):
        task_name = task['Task Name']
        recipient_email = task['Email']
        send_task_update_email(sender_email, recipient_email, task_name, row_index)

    return "Task update emails sent successfully!"

# Route to handle user feedback
@app.route('/feedback/<int:row_index>', methods=['GET', 'POST'])
def feedback(row_index):
    if request.method == 'POST':
        # Get the user's feedback from the form submission
        feedback = request.form['feedback']

        # Update the status in the Google Spreadsheet based on the feedback
        if feedback.lower() == 'done':
            update_status(row_index, 'Done')
        elif feedback.lower() == 'not done yet':
            update_status(row_index, 'Not Done Yet')

        # You can display a confirmation message to the user if desired.

    return render_template('feedback_template.html', row_index=row_index)

from flask import send_from_directory
@app.route('/static/<path:path>')
def serve_static(path):
    return send_from_directory('static', path)


if __name__ == '__main__':
    app.run()

