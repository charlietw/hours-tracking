import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

# for envvars
import os

# for emails
import smtplib
import email
from datetime import datetime

# Setting up gspread (see https://gspread.readthedocs.io/en/latest/)
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name(os.environ['HOURS_JSON_CREDENTIALS_DIR'], scope)

gc = gspread.authorize(credentials)

hrs = gc.open(os.environ['HOURS_SHEET_NAME']).sheet1

#TODO: 'Add headers' function
#TODO: Merge 'current_row' and 'last_reported_row' functions
#TODO: Split into two files? Functions + menu
#TODO: Fix bug when there is only one entry in a month
#TODO: Tidy 'month end' process i.e. updating the reported date

# Helper functions

def change_current_row(new_value):
    """
    Updates the 'current row' value
    """
    hrs.update_acell('C1', new_value)


def get_current_row():
    """
    Checks to see if a 'current row' value is stored, and sets it if not.
    Returns the value
    """
    current_value = hrs.acell('C1').value
    if current_value == "":
        current_value = 3
        change_current_row(current_value)
    else:
        current_value = int(current_value)
    return current_value


def change_last_reported_row(new_value):
    """
    Updates the 'last reported row' value
    """
    hrs.update_acell('C2', new_value)


def get_last_reported_row():
    """
    Checks to see if a 'current row' value is stored, and sets it if not.
    Returns the value
    """
    current_value = hrs.acell('C2').value
    if current_value == "":
        change_last_reported_row(3)
    else:
        current_value = int(current_value)
    return current_value


def minutes_to_hours(total_minutes):
    """
    Takes minutes and converts to hours/minutes
    """
    hours, minutes = divmod(total_minutes, 60)
    return hours, minutes

def send_email(text):
    """
    Sends an email with a payload in 'text'
    """
    gmail_user = os.environ['HOURS_EMAIL_ADDRESS']
    gmail_password = os.environ['HOURS_EMAIL_PASSWORD']
    try:
        msg = email.message.Message()
        msg['Subject'] = 'Hours'
        msg['From'] = gmail_user
        msg['To'] = gmail_user
        msg.add_header('Content-Type','text/html')
        msg.set_payload(text)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(gmail_user, gmail_password)
        server.sendmail(msg['From'], [msg['To']], msg.as_string())
        server.close()
    except Exception as e:
        print(e)


def print_current_row():
    """
    Takes the most recent row and prints details
    """
    current_row = get_current_row()
    values_list = hrs.row_values(current_row)
    try:
        date = values_list[0]
    except IndexError:
        date = "Empty"

    try:
        start_time = values_list[1]
    except IndexError:
        start_time = "Empty"

    try:
        end_time = values_list[2]
    except IndexError:
        end_time = "Empty"

    try:
        reason = values_list[6]
    except IndexError:
        reason = "Empty"

    row_description = "Row {0}. Date: {1}. Start time: {2}. End time: {3}. Reason: {4}".format(
        current_row,
        date,
        start_time,
        end_time,
        reason,
        )
    return row_description


def get_hours_report_data():
    """
    From the current row and the last reported row, returns
    the hours worked
    """
    current = get_current_row()
    last_reported = get_last_reported_row()
    report_data = []
    next_row = last_reported
    while next_row < current:
        report_data.append(hrs.row_values(next_row))
        next_row += 1
    return report_data


def get_total_hours():
    """
    Returns the total hours worked from the report range
    """
    report_range = hrs.range(
        "D{0}:D{1}".format(
            get_last_reported_row(),
            get_current_row() - 1
            )
        )
    total_minutes = 0
    for t in report_range:
        total_minutes += int(t.value)
        print(total_minutes)
    hours, minutes = minutes_to_hours(total_minutes)
    return hours, minutes


# Menu functions

def add_new_row():
    current_row = get_current_row()
    date = input("Enter the date (format DD/MM/YYYY): \n")
    start_time = input("Enter the START time (format HH:MM): \n")

    hrs.update_acell('A{0}'.format(current_row), date)
    hrs.update_acell('B{0}'.format(current_row), start_time)
    print("Row added, returning to menu...")
    menu()


def finish_current_row():
    current_row = get_current_row()
    end_time = input("Enter the END time (format HH:MM): \n")
    reason = input("Enter the reason for these hours: \n")
    time_format = "%H:%M"
    start_time = hrs.acell('B{0}'.format(current_row)).value
    # Convert the strings to timestamps
    start_time_formatted = datetime.strptime(start_time, time_format)
    end_time_formatted = datetime.strptime(end_time, time_format)
    time_diff = end_time_formatted - start_time_formatted # get the timedelta
    time_diff_seconds = time_diff.total_seconds() # convert to seconds
    total_minutes = int(divmod(time_diff_seconds, 60)[0]) # convert to minutes
    hrs.update_acell('C{0}'.format(current_row), end_time)
    hrs.update_acell('D{0}'.format(current_row), total_minutes)
    hours, minutes = minutes_to_hours(total_minutes)
    hrs.update_acell('E{0}'.format(current_row), hours)
    hrs.update_acell('F{0}'.format(current_row), minutes)
    hrs.update_acell('G{0}'.format(current_row), reason)
    print("{0} hours, {1} minutes recorded.".format(hours, minutes))
    change_current_row(current_row + 1)
    menu()


def change_active_row():
    new_row = int(input("Which row would you like to change to?\n"))
    change_current_row(new_row)
    print("Row changed to {0}, returning to menu...".format(new_row))
    menu()


def total_hours():
    print(get_total_hours())
    menu()


def email_hours():
    hours, minutes = get_total_hours()
    now = datetime.now()
    print("Retrieving hours...")
    report_data = get_hours_report_data()
    email_html = ""
    for r in report_data:
        email_html += "<tr><td>"
        email_html += r[0]
        email_html += "</td><td>"
        email_html += r[1]
        email_html += "</td><td>"
        email_html += r[2]
        email_html += "</td><td>"
        email_html += r[4]
        email_html += "</td><td>"
        email_html += r[5]
        email_html += "</td><td>"
        email_html += r[6]
        email_html += "</td></tr>"

    send_email("""
        <html>
            <body>
                <table border="1">
                    <tr>
                        <th>
                            Date
                        </th>
                        <th>
                            Start time
                        </th>
                        <th>
                            End time
                        </th>
                        <th>
                            Hours
                        </th>
                        <th>
                            Minutes
                        </th>
                        <th>
                            Reason
                        </th>
                    </tr>
                    {0}
                </table>
                <p><b>Total: {1} hours, {2} minutes </b></p>
                <p>
                    Signed: <u>Charlie Wren </u><br>
                    Date: <u>{3} </u>
                </p>
                <p>
                    Signed: __________________<br><br>
                    Date: __________________
                </p>

            </body>
        </html>
        """.format(
            email_html,
            hours,
            minutes,
            now.strftime("%d/%m/%Y")
            )
        )
    menu()



def menu():
    """
    Prints menu out, creating CLI
    """
    print("Latest entry: {0}".format(print_current_row()))
    print("1: Add a new entry")
    print("2: Finish the current entry")
    print("3: Change active row")
    print("4: Summary of hours")
    print("5: Email hours")
    print("0: Quit")
    user_input = int(input("Choose a menu option e.g. 1: \n"))
    if user_input == 1:
        add_new_row()
    if user_input == 2:
        finish_current_row()
    if user_input == 3:
        change_active_row()
    if user_input == 4:
        total_hours()
    if user_input == 5:
        email_hours()

menu()
