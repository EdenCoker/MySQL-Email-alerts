import smtplib
import mysql.connector
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
import time
import pandas as pd
import PySimpleGUI as sg

sg.theme('DarkBlue')

layout = [
    [sg.Text('MySQL Host:'), sg.Input(key='-MYSQL_HOST-')],
    [sg.Text('MySQL Username:'), sg.Input(key='-MYSQL_USER-')],
    [sg.Text('MySQL Password:'), sg.Input(key='-MYSQL_PASS-', password_char='*')],
    [sg.Text('MySQL Database:'), sg.Input(key='-MYSQL_DB-')],
    [sg.Text('Sender Email:'), sg.Input(key='-SENDER_EMAIL-')],
    [sg.Text('Sender Email Password:'), sg.Input(key='-EMAIL_PASS-', password_char='*')],
    [sg.Text('Recipient Email:'), sg.Input(key='-RECIPIENT_EMAIL-')],
    [sg.Text('Email Subject:'), sg.Input(key='-EMAIL_SUBJECT-')],
    [sg.Text('Email Message:'), sg.Input(key='-EMAIL_MESSAGE-')],
    [sg.Text('Excel File Name:'), sg.Input(key='-EXCEL_FILENAME-')],
    [sg.Text('Alert Table Name:'), sg.Input(key='-ALERT_TABLE-')],
    [sg.Text('Alert Column Name:'), sg.Input(key='-ALERT_COLUMN-')],
    [sg.Text('Target Table Name:'), sg.Input(key='-TARGET_TABLE-')],
    [sg.Text('Target Column Name:'), sg.Input(key='-TARGET_COLUMN-')],
    [sg.Button('Start'), sg.Button('Cancel')]
]

window = sg.Window('Email Alert System', layout)

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Cancel'):
        break
    if event == 'Start':
        mysql_hostid = values['-MYSQL_HOST-']
        mysql_userid = values['-MYSQL_USER-']
        mysql_password = values['-MYSQL_PASS-']
        mysql_dbname = values['-MYSQL_DB-']
        sender_email = values['-SENDER_EMAIL-']
        email_password = values['-EMAIL_PASS-']
        receiver_email = values['-RECIPIENT_EMAIL-']
        email_subject = values['-EMAIL_SUBJECT-']
        email_message = values['-EMAIL_MESSAGE-']
        Excel_file_name = values['-EXCEL_FILENAME-']
        Alert_table = values['-ALERT_TABLE-']
        Alert_column = values['-ALERT_COLUMN-']
        target_table = values['-TARGET_TABLE-']
        target_column = values['-TARGET_COLUMN-']

        mydb = mysql.connector.connect(
            host=mysql_hostid,
            user=mysql_userid,
            password=mysql_password,
            database=mysql_dbname
        )

        while True:
            today = datetime.now()
            # Check for new rows in my_log_table less than 6 hours ago
            query = f"SELECT COUNT(*) FROM {Alert_table} WHERE {Alert_column} >= %s"
            six_hours_ago = datetime.now() - timedelta(hours=6)
            values = (six_hours_ago,)
            cursor = mydb.cursor()
            cursor.execute(query, values)
            result = cursor.fetchone()
            # If there are new rows, send an email alert
            if result[0] > 0:
                # Get the data of the x most recent rows in the table tq_invest
                x = result[0]
                query = f"SELECT * FROM {target_table} ORDER BY {target_column} DESC LIMIT %s"
                values = (x,)
                cursor.execute(query, values)
                rows = cursor.fetchall()

                # Create DataFrame and replace "No Tax" values with "interest"
                df = pd.DataFrame(rows, columns=[i[0] for i in cursor.description])

                # Save DataFrame to Excel file
                df.to_excel(f'{Excel_file_name}{today.strftime("%Y-%m-%d")}.xlsx', index=False)

                # Compose email message
                msg = MIMEMultipart()
                msg['In-Reply-To'] = None
                msg['References'] = None
                msg['From'] = sender_email
                msg['To'] = receiver_email
                msg['Subject'] = email_subject
                body = email_message
                msg.attach(MIMEText(body, 'plain'))

                # Attach Excel file to email
                with open(f'{Excel_file_name}{today.strftime("%Y-%m-%d")}.xlsx', 'rb') as f:
                    attachment = MIMEApplication(f.read(), _subtype='xlsx')
                    attachment.add_header('Content-Disposition', 'attachment',
                                          filename=f'{Excel_file_name}{today.strftime("%Y-%m-%d")}.xlsx')
                    msg.attach(attachment)

                # Send email
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(sender_email, email_password)
                text = msg.as_string()
                server.sendmail(sender_email, receiver_email, text)
                server.quit()

                # Clear data in my_log_table
                query = f"TRUNCATE TABLE {Alert_table}"
                cursor.execute(query)
                mydb.commit()

            # Close database connection and wait for 60 seconds
            cursor.close()
            mydb.close()
            time.sleep(60)
            mydb = mysql.connector.connect(
                host=mysql_hostid,
                user=mysql_userid,
                password=mysql_password,
                database=mysql_dbname
            )
