"""
This module could be used to automate modification of Excel sheets
specifically is useful when you are dealing with excel files containing several sheets that are similar
and all sheets are needed to updated similarly

Working functions:
    - updating values of specific column in all sheet
    - updating column titles of all the sheets

"""
import openpyxl as xl


def modify_excel_values(name: "Excel file name", col: int, correction: int):
    """ Modifies all values of a specific column in Excel sheets

    Parameters:
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

    Parameters:
        name -- Excel file name
        new_titles -- list of new titles
    """
    print(f"*** Excel file: {name} ***")
    wb = xl.load_workbook(name)
    for sheet in wb.worksheets:
        print(f"working on {sheet}")
        for col in range(1, len(new_titles)+1):
            current_title = sheet.cell(1, col).value
            updated_title = new_titles[col-1]
            sheet.cell(1, col).value = updated_title
            print(f"{current_title} --> {updated_title}")

    wb.save(name)
