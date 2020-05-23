
# for envvars
import os

import argparse

from functions import *


#TODO: 'Add headers' function
#TODO: Merge 'current_row' and 'last_reported_row' functions
#TODO: Fix bug when there is only one entry in a month
#TODO: Tidy 'month end' process i.e. updating the reported date



workbook, worksheet_list, hrs_sheet, months_sheet = gspread_setup()


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
    print("6: Email hours & finalise month (BROKEN)")
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
    elif user_input == 6:
        # email_hours(hrs_sheet)
        month_end(hrs_sheet, months_sheet)
        menu()

# Adding CLI args
parser = argparse.ArgumentParser()

parser.add_argument(
    "-i", "--interactive",
    help="run the interactive CLI",
    action="store_true")

parser.add_argument(
    "-e", "--email",
    help="send the hours you have recorded in the month",
    action="store_true")

parser.add_argument(
    "-s", "--setup",
    help="initial setup of worksheets",
    action="store_true")

args = parser.parse_args()

if args.interactive:
    menu()

if args.email:
    email_hours(hrs_sheet)

if args.setup:
    # Doing initial setup to check that the sheets are set up properly
    hrs_sheet = False
    months_sheet = False

    for w in worksheet_list:
        if w.title == "Hours":
            print("Hours worksheet found")
            hrs_sheet = workbook.worksheet("Hours")
        if w.title == "Months":
            print("Months worksheet found")
            months_sheet = workbook.worksheet("Months")

    if hrs_sheet == False:
        print("WARNING: Create a worksheet named 'Hours' before continuing")
    if months_sheet == False:
        print("WARNING: Create a worksheet named 'Months' before continuing")

