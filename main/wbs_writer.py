"""WBS Excel file generation and download.

Functions for creating WBS Excel files from AI-generated markdown tables
and copying results to the user's Downloads folder.
No GUI dependencies — callers are responsible for error display.
"""

import io
import os
import re
import shutil
from datetime import datetime

import pandas as pd
import pywintypes
import win32com.client as win32
from openpyxl.utils.dataframe import dataframe_to_rows


def markdown_table_to_dataframe(content: str) -> pd.DataFrame:
    """Extract a markdown table from text content and convert to DataFrame.

    Args:
        content: Text containing a markdown table with pipe-delimited rows.

    Returns:
        A cleaned DataFrame with proper column headers.
    """
    table_pattern = re.compile(r'\|.*\|')
    markdown_table = '\n'.join(table_pattern.findall(content))

    data = io.StringIO(markdown_table)
    df = pd.read_csv(data, sep="|", skipinitialspace=True, engine='python')

    # Remove the separator row
    df = df.iloc[1:]

    # Drop the first and last columns (empty from pipe format)
    df = df.drop(df.columns[[0, -1]], axis=1)

    # Use first data row as new column headers
    df = df.rename(columns=df.iloc[0]).drop(df.index[0])

    return df


def write_wbs_to_excel(df: pd.DataFrame, start_date, end_date) -> None:
    """Write WBS DataFrame to an Excel template using COM automation.

    Args:
        df: The WBS data as a DataFrame.
        start_date: Project start date (date object).
        end_date: Project end date (date object).

    Raises:
        FileNotFoundError: If the template file is not found.
        PermissionError: If the file can't be written.
        Exception: For any other Excel COM errors.
    """
    template_path = 'JDU-WBS_Template_Samples.xlsm'
    macro_name = 'UpdateDatesAndFormat'
    excel = win32.gencache.EnsureDispatch("Excel.Application")

    # Open the workbook
    workbook = excel.Workbooks.Open(os.path.abspath(template_path))
    sheet = workbook.Sheets(1)

    # Write metadata
    sheet.Cells(2, 2).Value = "Details_WBS.xlsm"
    sheet.Cells(6, 2).Value = start_date.strftime('%m/%d/%Y')
    sheet.Cells(7, 2).Value = end_date.strftime('%m/%d/%Y')

    # Set current date
    current_date = datetime.now().date()
    pywintypes_time = pywintypes.Time(current_date)
    sheet.Cells(2, 7).Value = pywintypes_time

    # Write the DataFrame to the Excel template starting at row 10
    start_row = 10
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
        for c_idx, value in enumerate(row, 1):
            sheet.Cells(r_idx, c_idx).Value = value

    # Run the macro
    excel.Application.Run(macro_name)

    # Save and close the workbook
    workbook.SaveAs(os.path.abspath("Details_WBS.xlsm"))
    workbook.Close()

    # Quit Excel
    excel.Application.Quit()

    print("DataFrame saved to Details_WBS.xlsm")


def copy_to_downloads() -> str:
    """Copy the generated WBS file to the user's Downloads folder.

    Returns:
        The destination file path.

    Raises:
        Exception: If the copy operation fails.
    """
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    destination_file_path = os.path.join(downloads_folder, "Details_WBS.xlsm")

    # Create dummy file in the download folder
    df = pd.DataFrame()
    df.to_excel(destination_file_path, index=False)

    # Get the current directory and copy
    current_directory = os.getcwd()
    source_file_path = os.path.join(current_directory, "Details_WBS.xlsm")

    shutil.copy(source_file_path, destination_file_path)
    return destination_file_path
