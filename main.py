import re
from openpyxl import Workbook
import win32com.client
import os
import pythoncom  # Import pythoncom library
import xlwings as xw
import pandas as pd
import numpy as np

folder_path = r"M:\CDB\Swift\AllianceLiteMessages"
output_folder = r"M:\CDB\Swift\AllianceLiteMessages\Processed Excels"

def inject_macro(excel_file_path, macro_name, file_base_name):
    print(file_base_name)
    macro_code = f'''
                    Sub getData()
                    '
                    ' getData Macro
                    ' A macro to get data from the Swift files
                    '
                    ' Keyboard Shortcut: Ctrl+Shift+D
                    '
                        ActiveWorkbook.Queries.Add Name:="{file_base_name}", Formula:= _
                            "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(File.Contents(""M:\\CDB\\Swift\\AllianceLiteMessages\\{file_base_name}.txt""),[Delimiter="":"", Columns=2, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{{{""Column1"", type text}}, {{""Column2"", type text}}}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
                        ActiveWorkbook.Worksheets.Add
                        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""{file_base_name}"";Extended Properties=""""" _
                            , Destination:=Range("$A$1")).QueryTable
                            .CommandType = xlCmdSql
                            .CommandText = Array("SELECT * FROM [{file_base_name}]")
                            .RowNumbers = False
                            .FillAdjacentFormulas = False
                            .PreserveFormatting = True
                            .RefreshOnFileOpen = False
                            .BackgroundQuery = True
                            .RefreshStyle = xlInsertDeleteCells
                            .SavePassword = False
                            .SaveData = True
                            .AdjustColumnWidth = True
                            .RefreshPeriod = 0
                            .PreserveColumnInfo = True
                            .ListObject.DisplayName = "_{file_base_name}"
                            .Refresh BackgroundQuery:=False
                        End With
                    End Sub
            '''

    pythoncom.CoInitialize()

    com_instance = win32com.client.Dispatch("Excel.Application")
    com_instance.Visible = False
    com_instance.DisplayAlerts = False

    workbook = com_instance.Workbooks.Add()
    xlmodule = workbook.VBProject.VBComponents.Add(1)
    xlmodule.Name = macro_name  # Set the name of the module
    xlmodule.CodeModule.AddFromString(macro_code)
    workbook.SaveAs(Filename=excel_file_path, FileFormat=52)
    workbook.Close()
    com_instance.Quit()

    pythoncom.CoUninitialize()

def execute_macro(file_path):
    app = xw.App(visible=False)  # Ensure Excel opens in the background
    wb = app.books.open(file_path)

    try:
        macro_name = "getData"
        excel_macro = wb.macro(macro_name)
        excel_macro()  # Execute the macro
    except Exception as e:
        print(f"Error occurred while running macro: {e}")
    finally:
        # Save and close the workbook
        wb.save()
        wb.close()
        app.quit()

def print_first_account_holder(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Assuming column A is named 'A' and column B is named 'B'. Adjust as necessary.
    column_a = 'Column1'  # Adjust if your DataFrame uses a different label
    column_b = 'Column2'  # Adjust if your DataFrame uses a different label

    # Initialize an empty DataFrame for the results
    results_df = pd.DataFrame()

    # Track the current account holder and the keys found for it
    current_account_holder = None
    keys_found = set()

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        # Check for 'AccountHolder' and update the current account holder
        if row[column_a] == 'AccountHolder':
            current_account_holder = row[column_b]  # Assuming the account holder's name is in column B
            keys_found.clear()  # Reset keys for the new account holder
        else:
            # Process the key-value pairs
            key = row[column_a]
            value = row[column_b]

            # If it's a new key, add a new column for it
            if key not in keys_found:
                keys_found.add(key)
                if key not in results_df.columns:
                    results_df[key] = np.nan  # Initialize the column with NaNs
                results_df.at[current_account_holder, key] = value  # Add the value for the current account holder
            else:
                # If the key exists, find the first empty cell in its column for the current account holder
                if current_account_holder in results_df.index:
                    next_row = results_df[key].isnull().idxmax()  # Get the next available row
                    if pd.isnull(results_df.at[next_row, key]):
                        results_df.at[next_row, key] = value  # Insert the value if the cell is empty
                    else:
                        # If no empty cell is available, append a new row
                        new_row = pd.Series(index=[current_account_holder], data={key: value})
                        results_df = results_df.append(new_row, ignore_index=False)
                else:
                    # If the account holder does not exist in the index, simply add the value
                    results_df.at[current_account_holder, key] = value

    results_df.fillna(value=pd.NA, inplace=True)

    print(results_df)


if os.path.exists(folder_path) and os.path.isdir(folder_path):
    files = os.listdir(folder_path)

    print("Text files in the directory:")
    for file_name in files:
        if file_name.endswith(".txt"):
            file_path = os.path.join(folder_path, file_name)
            file_base_name = os.path.splitext(file_name)[0]

            with open(file_path, 'r') as file:
                for line in file:
                    match = re.search(r'ISIN:\s*(\w{12})', line)
                    if match:
                        isin = match.group(1)
                        print(f"ISIN: {isin}")
                        wb = Workbook()
                        ws = wb.active
                        output_file = os.path.join(output_folder, f"{isin}_{file_base_name}.xlsm")
                        wb.save(output_file)
                        print(f"Excel file saved for ISIN {isin} at: {output_file}")

                        inject_macro(output_file, 'loader', file_base_name)
                        execute_macro(output_file)
                        print_first_account_holder(output_file)

else:
    print("The specified directory does not exist or is not a directory.")