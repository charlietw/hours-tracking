# for emails
import smtplib
import email

# for envvars
import os

# For getting current date, and timedelta
from datetime import datetime

# Helper functions

def change_current_row(sheet, new_value):
    """
    Updates the 'current row' value
    """
    sheet.update_acell('C1', new_value)


def get_current_row(sheet):
    """
    Checks to see if a 'current row' value is stored, and sets it if not.
    Returns the value
    """
    current_value = sheet.acell('C1').value
    if current_value == "":
        current_value = 4
        change_current_row(sheet, current_value)
    else:
        current_value = int(current_value)
    return current_value


def change_last_reported_row(sheet, new_value):
    """
    Updates the 'last reported row' value
    """
    sheet.update_acell('C2', new_value)


def get_last_reported_row(sheet):
    """
    Checks to see if a 'current row' value is stored, and sets it if not.
    Returns the value
    """
    current_value = sheet.acell('C2').value
    if current_value == "":
        change_last_reported_row(4)
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


def print_current_row(sheet):
    """
    Takes the most recent row and prints details
    """
    current_row = get_current_row(sheet)
    values_list = sheet.row_values(current_row)
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


def get_hours_report_data(sheet):
    """
    From the current row and the last reported row, returns
    the hours worked
    """
    current = get_current_row(sheet)
    last_reported = get_last_reported_row(sheet)
    report_data = []
    next_row = last_reported
    while next_row < current:
        report_data.append(sheet.row_values(next_row))
        next_row += 1
    return report_data


def get_total_hours(sheet):
    """
    Returns the total hours worked from the report range
    """
    report_range = sheet.range(
        "D{0}:D{1}".format(
            get_last_reported_row(sheet),
            get_current_row(sheet) - 1
            )
        )
    total_minutes = 0
    for t in report_range:
        total_minutes += int(t.value)
        print(total_minutes)
    hours, minutes = minutes_to_hours(total_minutes)
    return hours, minutes


# Menu functions

def add_new_row(sheet):
    current_row = get_current_row(sheet)
    date = input("Enter the date (format DD/MM/YYYY): \n")
    start_time = input("Enter the START time (format HH:MM): \n")

    sheet.update_acell('A{0}'.format(current_row), date)
    sheet.update_acell('B{0}'.format(current_row), start_time)
    print("Row added, returning to menu...")


def finish_current_row(sheet):
    current_row = get_current_row(sheet)
    end_time = input("Enter the END time (format HH:MM): \n")
    reason = input("Enter the reason for these hours: \n")
    time_format = "%H:%M"
    start_time = sheet.acell('B{0}'.format(current_row)).value
    # Convert the strings to timestamps
    start_time_formatted = datetime.strptime(start_time, time_format)
    end_time_formatted = datetime.strptime(end_time, time_format)
    time_diff = end_time_formatted - start_time_formatted # get the timedelta
    time_diff_seconds = time_diff.total_seconds() # convert to seconds
    total_minutes = int(divmod(time_diff_seconds, 60)[0]) # convert to minutes
    sheet.update_acell('C{0}'.format(current_row), end_time)
    sheet.update_acell('D{0}'.format(current_row), total_minutes)
    hours, minutes = minutes_to_hours(total_minutes)
    sheet.update_acell('E{0}'.format(current_row), hours)
    sheet.update_acell('F{0}'.format(current_row), minutes)
    sheet.update_acell('G{0}'.format(current_row), reason)
    print("{0} hours, {1} minutes recorded.".format(hours, minutes))
    change_current_row(sheet, current_row + 1)


def change_active_row(sheet):
    new_row = int(input("Which row would you like to change to?\n"))
    change_current_row(sheet, new_row)
    print("Row changed to {0}, returning to menu...".format(new_row))


def total_hours(sheet):
    print(get_total_hours(sheet))


def email_hours(sheet):
    hours, minutes = get_total_hours(sheet)
    now = datetime.now()
    print("Retrieving hours...")
    report_data = get_hours_report_data(sheet)
    email_html = ""
    for r in report_data:
        email_html += "<tr><td>{0}</td>".format(r[0])
        email_html += "<td>{0}</td>".format(r[1])
        email_html += "<td>{0}</td>".format(r[2])
        email_html += "<td>{0}</td>".format(r[3])
        email_html += "<td>{0}</td>".format(r[4])
        email_html += "<td>{0}</td>".format(r[5])
        email_html += "<td>{0}</td></tr>".format(r[6])

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