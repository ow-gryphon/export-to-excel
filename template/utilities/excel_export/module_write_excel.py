

import xlwings as xw # Documentation https://docs.xlwings.org/en/stable/connect_to_workbook.html
import numpy as np 
import pandas as pd
import os
import shutil
from openpyxl.utils import get_column_letter, column_index_from_string
import re
import time
from datetime import datetime
# Automating document generation could be performed with the library Python-docx
import seaborn as sns # For importing a dummy dB

# This module writes the Excel file

class ExcelWriter():

    def __init__(self, config_Excel):
        # Include all the inputs that are needed from config modeule
        self.ID = config_Excel["ID"]
        self.template_folder = config_Excel["template_folder"]
        self.template_file_in = config_Excel["template_file"]
        self.target_folder = config_Excel["output_folder"]
        self.template_file_out = config_Excel["output_file"] 

        self.file_in = os.path.join(self.template_folder, self.template_file_in)
        self.file_out = os.path.join(self.target_folder, self.template_file_out)

        self.data_to_write = []


    def df_to_worksheet(self, df, worksheet, from_cell):
        """Copy-paste dataframe value (non NA) into specific worksheet range
        Args:
        df: DataFrame to copy from
        worksheet: Worksheet to paste in
        from_cell: top-left cell of the worksheet where you want to paste df
        col_limit: column number up to which overwrite worksheet cells by 0 beyond df size
        row_limit: row number up to which overwrite worksheet cells by 0 beyond df size
        Returns:
        worksheet with copy-pasted dataframe
        """
        cell = worksheet[from_cell]
        row_idx = cell.row
        col_idx = cell.column
        col_limit = df.shape[1]
        row_limit = df.shape[0]
        values = []
        for r in range(row_limit):
            row_values = []
            for c in range(col_limit):
                if pd.isna(df.index[r]):
                    value = np.nan
                else: # elif c < col_idx:
                    value = df.iat[r, c]
                # else:
                #     value = 0
                if not pd.notna(value):
                    value = worksheet.range((r + row_idx, c + col_idx)).value
                row_values.append(value)
            values.append(row_values)
        col_ini = column_index_from_string(re.sub(r'[^a-zA-Z]', '', from_cell))
        cell_range = f"{from_cell}:{get_column_letter(col_ini + col_limit -1)}{row_idx + row_limit-1}"
        print(f'- - - - - - Cell range: {cell_range}')
        worksheet.range(cell_range).value = values

    def print_error(self):
        print(" * * * E R R O R - I N - E X C E L - R E P O R T I N G * * * ")


    def prepare_output_path(self):
        if os.path.exists(self.file_out):
            try:
                os.remove(self.file_out)
                print("- - - EXCEL - A previous version of the template has been deleted")
                print("")
            except Exception as e:
                self.print_error()
                print("The Excel could not be deleted")
        elif not os.path.exists(self.target_folder):
            os.mkdir(self.target_folder)
            
        # shutil.copyfile(self.file_in, self.file_out)


    def dict_to_df(self, dictionary, columns=['Item', 'Value']):
        keys = [key.capitalize() for key in dictionary.keys()]
        values = dictionary.values()

        df = pd.DataFrame(values, keys).reset_index()
        df.columns = columns
        return df

    
    def append_df_to_write(self, df, worksheet_name, from_cell):
        df_names = pd.DataFrame(df.columns).transpose()
        (i_max, j_max) = df.shape
        colnames = [f'col{j}' for j in range(0,j_max) ]
        df_names.columns = colnames
        df.columns = colnames
        df = pd.concat([df_names, df], axis=0)

        dict_temp = {"df": df,
                    "worksheet_name": worksheet_name,
                    "from_cell": from_cell}

        self.data_to_write.append(dict_temp)

    def append_dict_to_write(self, dictionary, worksheet_name, from_cell):

        df = self.dict_to_df(dictionary)
        self.append_df_to_write(df, worksheet_name, from_cell)
         

    def write_Excel(self):
        n_max = 20
        for n in range(0,n_max):
            now = datetime.now()
            time_now = now.strftime("%H:%M:%S")
            print("")
            print(f"- - - {time_now} - - - EXCEL - Writing the Excel file - Attempt {n} out of {n_max} - - -")
            try: 
                with xw.App(visible=False) as app:
                    
                        print(f"- - - - - - Using the template {self.file_in}")
                        workbook =  app.books.open(self.file_in)

                        for val in self.data_to_write:

                            df = val["df"]
                            worksheet_name = val["worksheet_name"]
                            from_cell = val["from_cell"]
                            now = datetime.now()
                            time_now = now.strftime("%H:%M:%S")
                            print(f'- - - {time_now} - - - Sheet - {worksheet_name} | Cell: {from_cell}')
                            
                            
                            
                            worksheet = workbook.sheets[worksheet_name]

                            self.df_to_worksheet(df, worksheet, from_cell)

                        workbook.save(self.file_out)
                        workbook.close()

                        break
            except Exception as e:
                self.print_error()
                print(f"Error: {e}")

