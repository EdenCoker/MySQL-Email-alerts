# MySQL-Email-alerts


This is a Python program that creates an email alert system. The program connects to a MySQL database to monitor changes in a particular table and sends an email alert to a specified recipient when new rows are added to that table. The program allows the user to input their MySQL database credentials, email credentials, and various other parameters such as the name of the Excel file to be generated and the table and column names to monitor for changes.

The program uses the PySimpleGUI library to create a simple GUI with input fields for the user to enter the required parameters. Once the user clicks the "Start" button, the program connects to the MySQL database and begins monitoring the specified table for new rows. If new rows are detected, the program extracts the data from the x most recent rows and saves it to an Excel file with the current date in the file name. It then composes an email message with the specified subject and message and attaches the Excel file to the email before sending it to the specified recipient.

Finally, the program clears the data in the monitored table and waits for 60 seconds before checking for new rows again. The program continues to run indefinitely until the user clicks the "Cancel" button or closes the GUI window.

Note that this program assumes that the user has already set up a Gmail account to use as the sender email and has allowed less secure apps to access their account. If the user is using a different email provider, they may need to modify the program accordingly.
