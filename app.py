# Flask Setup
import json

import secret
import os
from datetime import datetime
from flask_mail import Mail, Message
from flask import Flask, jsonify, request, abort
# Google Sheets API Setup
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#region setup
EMAIL = 'user@gmail.com'
PROJECT_DICT = {"name_1": "jhon dupont", "name_2": "joel duron", "name_3": "julien monrau"}
app = Flask(__name__)
mail = Mail(app)
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = EMAIL
app.config['MAIL_PASSWORD'] = secret.pwd
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)
credential = ServiceAccountCredentials.from_json_keyfile_name("Keys.json",
                                                              ["https://spreadsheets.google.com/feeds",
                                                               "https://www.googleapis.com/auth/spreadsheets",
                                                               "https://www.googleapis.com/auth/drive.file",
                                                               "https://www.googleapis.com/auth/drive"])
client = gspread.authorize(credential)
google_sheet = client.open("Recruitment").sheet1
google_sheet_values = google_sheet.get_values()
f = open('mail_content.json')
mail_content = json.load(f)
#endregion

def assert_constraints():
    id_values = google_sheet.col_values(1)
    email_values = google_sheet.col_values(2)
    project_values = google_sheet.col_values(3)
    status_values = google_sheet.col_values(4)
    if "" in id_values:
        return "Error Id."
    if "" in email_values:
        return "Error Email."
    if '' in project_values:
        return "Error Project."
    if '' in status_values:
        return "Error Status."
    return False

def days_between(d1, today):
    try:
        d1 = datetime.strptime(d1,"%d/%m/%Y %X")
        return abs((today - d1).days)
    except:
        return 0

def sendOnlineTest(idx,row):
    # send email
    project = row[2]
    msg = Message(mail_content["onlineTest"]["subject"], sender=EMAIL, recipients=[row[1]])
    msg.body = mail_content["onlineTest"]["content"].format(project=project)
    mail.send(msg)
    # change status value to Online Test Sent
    google_sheet.update_cell(idx + 1, 4, "Online Test Sent")
    # change mail sent to today's dateTime
    date_time = datetime.now()
    google_sheet.update_cell(idx + 1, 5, date_time.strftime("%d/%m/%Y %X"))

def sendOnlineTestReminder(idx,row):
    # send reminder mail
    project = row[2]
    msg = Message(mail_content["reminderTest"]["subject"], sender=EMAIL, recipients=[row[1]])
    msg.body = mail_content["reminderTest"]["content"]
    mail.send(msg)
    # change status value to Online Test Sent
    google_sheet.update_cell(idx + 1, 4, "Reminder Sent")
    # change mail sent to today's dateTime
    date_time = datetime.now()
    google_sheet.update_cell(idx + 1, 5, date_time.strftime("%d/%m/%Y %X"))

def sendTestResult(idx,row):
    total = row[5]
    total_list = total.split('/')
    total_list = list(map(float, total_list))
    score = total_list[0]
    moyenne = total_list[1] / 2.0
    # change mail sent to today's dateTime
    date_time = datetime.now()
    google_sheet.update_cell(idx + 1, 5, date_time.strftime("%d/%m/%Y %X"))
    if score >= moyenne:
        # send reminder mail
        msg = Message(mail_content["interviewMail"]["subject"], sender=EMAIL, recipients=[row[1]])
        project = row[2]
        name = PROJECT_DICT[project]
        msg.body = mail_content["interviewMail"]["content"].format(name=name)
        mail.send(msg)
        # change status value to Online Test Sent
        google_sheet.update_cell(idx + 1, 4, "Interview Mail Sent")
    else:
        # send reminder mail
        msg = Message(mail_content["refusalMail"]["subject"], sender=EMAIL, recipients=[row[1]])
        msg.body = mail_content["refusalMail"]["content"]
        mail.send(msg)
        # change status value to Online Test Sent
        google_sheet.update_cell(idx + 1, 4, "Refusal Mail Sent")

def operations():
    for idx, row in enumerate(google_sheet_values):
        if row[3] == "Applied":
            sendOnlineTest(idx, row)
        elif (row[3] == "Online Test Sent") & (days_between(row[4], datetime.now()) >= 7) & (row[5] == ''):
            sendOnlineTestReminder(idx, row)
        elif row[3] == "Submitted Test":
            sendTestResult(idx, row)

@app.route('/')
def index():
    try:
        check = assert_constraints()
        if check:
            return f"<h1>{check}</h1>"
        else:
            operations()
            return "<h1>Operations Done With Success.</h1>"
    except Exception as e:
        print(e)
        return"<h1>Something went wrong</h1>"


if __name__ == "__main__":
    app.run(host='0.0.0.0', debug=False, port=os.environ.get('PORT', 80))