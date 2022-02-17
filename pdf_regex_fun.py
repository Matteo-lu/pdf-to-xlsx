#!/usr/bin/env python3

"""spreadsheet auto-fill auxiliar functions"""


import re
import ctypes

def pdf_regex(pdf_text, recent_pdf_file):
    """Function to extract information from the pdf text variable
    using regular expressions.

    Args:
        pdf_text: variable containing the text extracted from the pdf file
        recent_pdf_file: File name

    Returns:
        The return value. dictionary containing the values to be
        added to the .xlsx file

    """

    MB_OK = 0x0
    ICON_EXLAIM=0x30
    MessageBox = ctypes.windll.user32.MessageBoxW

    dict_values = {}
    # Getting the quotation number
    quotation_number = re.compile(r'No. *\d\d\d\d')
    mo = quotation_number.search(pdf_text)
    if not mo:
        MessageBox(
            None,
            "The quote number couldn't be found inside the file " + recent_pdf_file,
            'No data',
            MB_OK | ICON_EXLAIM
            )
        return(None)
    dict_values['quot_number'] = mo.group()

    # Getting the expedition date
    dates = re.compile(r'(\d\d\W\d\d\W\d\d\d\d)+')
    mo = dates.search(pdf_text)
    if not mo:
        MessageBox(
            None,
            "The expedition date couldn't be found inside the file " + recent_pdf_file,
            'No data',
            MB_OK | ICON_EXLAIM
            )
        return(None)
    dates_string = mo.group()
    dates_list = []
    dates_list.append(dates_string[:10])
    dates_list.append(dates_string[10:])
    dict_values['expedition_date'] = dates_list[0]

    # Responsable
    dict_values['responsable'] = 'Lina García'

    # Getting cliente
    client_list = re.compile(r'\d{9}\D+')
    mo = client_list.search(pdf_text)
    if not mo:
        MessageBox(
            None,
            "The client couldn't be found inside the file " + recent_pdf_file,
            'No data',
            MB_OK | ICON_EXLAIM
            )
        return(None)
    dict_values['client'] = mo.group()[9:]

    # Value before IVA
    value_list = re.compile(
        r'Subtotal\$(\d+,)?(\d+,)?(\d+,)?(\d+,)?(\d+.)?(\d+)?'
    )
    mo = value_list.search(pdf_text)
    if not mo:
        MessageBox(
            None,
            "The Value before IVA couldn't be found inside the file " + recent_pdf_file,
            'No data',
            MB_OK | ICON_EXLAIM
            )
        return(None)
    dict_values['subtotal'] = mo.group()[9:]

    return(dict_values)
