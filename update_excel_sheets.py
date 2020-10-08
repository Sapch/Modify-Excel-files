"""
This module could be used to automate modification of Excel sheets
TODO:
complete functions:
    - updating column titles of all the sheets

Working functions:
    - updating values of specific column in all sheet

"""
import openpyxl as xl


def modify_excel_values(name: "Excel file name", col: int, correction: int):
    """ Modifies all values of a specific column in Excel sheets

    Arguments:
        name -- Excel file name
        col -- col number to be corrected
        correction -- a correction value between 0 to 1
    """
    print(f"*** Excel file: {name} ***")
    wb = xl.load_workbook(name)
    for sheet in wb.worksheets:
        print(f"working on {sheet}")
        for row in range(2, sheet.max_row + 1):
            current_value = sheet.cell(row, col).value
            corrected_value = current_value * correction
            sheet.cell(row, col).value = corrected_value
            print(f"row{row}: {current_value} --> {corrected_value}")

    wb.save(name)


def modify_excel_sheet_col_titles(name: "Excel file name", new_titles: list):
    """ Modifies titles of columns in all of the sheets

    Arguments:
        name -- Excel file name
        new_titles -- list of new titles
    """