from typing import List

import pandas as pd
from openpyxl.styles import Alignment

"""
This class is the processor to do analysis of the data
and export the result to .xlsx format file. 
"""
class Processor:

    """
    This method is to save the data to excel.
    """
    def save_table_as_excel(self, column_headers: List[str], rows_data: List[str], filename: str, sheet_name: str, append_mode: bool=False) -> None:

        # Convert the data into a pandas DataFrame
        df = pd.DataFrame(rows_data, columns=column_headers)

        # Ensure the filename ends with '.xlsx'
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'

        # Append to an existing Excel file or create a new one
        if append_mode:
            with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                workbook = writer.book
                worksheet = workbook[sheet_name]
        else:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                workbook = writer.book
                worksheet = workbook[sheet_name]

        # Apply alignment to each cell in the worksheet
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # Save the workbook
        workbook.save(filename)