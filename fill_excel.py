"""Module containing th fill excel function"""


from openpyxl.styles import PatternFill, Border, Side
from openpyxl.styles import Alignment, Protection, Font
from datetime import datetime


def fill_excel(dict_values, work_sheet, empty_cell):
    """Function to apply style and insert values to each cell from A# to Q#

    Args:
        dict_values: dictionary containing the values to be
        added to the .xlsx file
        work_sheet: .xlsx worksheet object
        empty_cell: first empty cell identify in the .xlsx file

    Returns:
        The return value. None
    """

    # Setting variables fill and align with the expecting cells format
    fll = PatternFill(start_color="a4c2f4", fill_type="solid")
    align = Alignment(horizontal='center', vertical='center', wrapText=True)

    letter = 'ABCDEFGHIJKLMNOPQ'
    for cell in letter:
        if cell == 'A':
            work_sheet[cell + str(empty_cell)] = dict_values['quot_number']
        if cell == 'C':
            date_format = datetime.strptime(
                dict_values['expedition_date'], '%d/%m/%Y')
            work_sheet[cell + str(empty_cell)] = date_format.date()
        if cell == 'F':
            work_sheet[cell + str(empty_cell)] = dict_values['responsable']
        if cell == 'H':
            work_sheet[cell + str(empty_cell)] = dict_values['client']
        if cell == 'J':
            work_sheet[cell + str(empty_cell)] = dict_values['subtotal']

        style_cell = work_sheet[cell + str(empty_cell)]
        style_cell.fill = fll
        style_cell.alignment = align
