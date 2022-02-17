"""Module containing th recent_pdf function"""


import glob
import os


def recent_pdf():
    """Function to get the most recent PDF file from the current directory

    Args:
        None

    Returns:
        The return value. Most recent file name within the current directory
    """

    file_path = '*.pdf'
    list_files = sorted(glob.iglob(file_path),
                        key=os.path.getctime, reverse=True)
    if (len(list_files) == 0):
        return(None)
    return (list_files)
