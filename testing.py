import unittest
from functions import *


class TestChangeCurrentRow(unittest.TestCase):
    """
    POC for test cases
    """
    workbook, worksheet_list, hrs_sheet, months_sheet = gspread_setup()

    def setUp(self):
        self.current_row_original = get_current_row(self.hrs_sheet)
        print(self.current_row_original)

    def test_change_row(self):
        change_current_row(self.hrs_sheet, 1) # changes current row to 1
        self.assertEqual(get_current_row(self.hrs_sheet), 1)

    def tearDown(self):
        change_current_row(self.hrs_sheet, self.current_row_original) # changes current row back


if __name__ == '__main__':
    unittest.main()