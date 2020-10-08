"""
This module could be used to automate modification of Excel sheets
TODO:
complete functions:
    - updating values of specific column in all sheet
    - updating column titles of all the sheets

"""

import openpyxl as xl

def modify_excel_values(name: "Excel file name", col: int, correction: int):
    """ Modifies existing values of a specific column in Excel sheets

    Arguments:
        name -- Excel file name
        col -- col number to be corrected
        correction -- a correction value between 0 to 1
    """

def modify_excel_sheet_col_titles(name: "Excel file name", new_titles: list):
    """ Modifies titles of columns in all of the sheets

    Arguments:
        name -- Excel file name
        new_titles -- list of new titles
    """