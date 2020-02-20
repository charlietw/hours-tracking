import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

# for envvars
import os

from functions import *

# Setting up gspread (see https://gspread.readthedocs.io/en/latest/)
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name(os.environ['HOURS_JSON_CREDENTIALS_DIR'], scope)

gc = gspread.authorize(credentials)

hrs_sheet = gc.open(os.environ['HOURS_SHEET_NAME']).sheet1

#TODO: 'Add headers' function
#TODO: Merge 'current_row' and 'last_reported_row' functions
#TODO: Fix bug when there is only one entry in a month
#TODO: Tidy 'month end' process i.e. updating the reported date


def menu():
    """
    Prints menu out, creating CLI
    """
    print("Latest entry: {0}".format(print_current_row(hrs_sheet)))
    print("1: Add a new entry")
    print("2: Finish the current entry")
    print("3: Change active row")
    print("4: Summary of hours")
    print("5: Email hours")
    print("0: Quit")
    user_input = int(input("Choose a menu option e.g. 1: \n"))
    if user_input == 1:
        add_new_row(hrs_sheet)
        menu()
    elif user_input == 2:
        finish_current_row(hrs_sheet)
        menu()
    elif user_input == 3:
        change_active_row(hrs_sheet)
        menu()
    elif user_input == 4:
        total_hours(hrs_sheet)
        menu()
    elif user_input == 5:
        email_hours(hrs_sheet)
        menu()

menu()
