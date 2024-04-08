import re
import win32com.client
import os
import pythoncom
import xlwings as xw
import psutil
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

folder_path = r"M:\CDB\Swift\AllianceLiteMessages\Messages"
output_folder = r"M:\CDB\Swift\AllianceLiteMessages\Processed Excels"

def kill_excel_processes():
    for process in psutil.process_iter(['pid', 'name']):
        # Check if the process name is Excel
        if 'EXCEL.EXE' in process.info['name'].upper():
            print(f"Terminating Excel process {process.info['pid']}...")
            psutil.Process(process.info['pid']).kill()

def inject_macro(macro_name, file_base_name, output_file):
    print('File name is : ' + file_base_name)
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

    try:
        # Call this function before initializing Excel with win32com
        kill_excel_processes()

        # Proceed with your existing code to inject macros
        pythoncom.CoInitialize()
        excel_app = win32com.client.Dispatch("Excel.Application")
        # Make sure to set Visible to True if you want to see the Excel application
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Add()

        # Adding the macro to the workbook
        xlmodule = workbook.VBProject.VBComponents.Add(1)  # 1 corresponds to vbext_ct_StdModule, a standard module
        xlmodule.Name = macro_name
        xlmodule.CodeModule.AddFromString(macro_code)

        # Save the workbook with the injected macro at the specified path
        workbook.SaveAs(Filename=output_file, FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        workbook.Close(False)
        excel_app.Quit()
        pythoncom.CoUninitialize()
        print(f"Workbook saved successfully at: {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
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
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path)
    print(df.head())

    # Dictionary to hold the {key: value} pairs
    account_holder_data = {}
    current_key = None

    start_appending = False

    for index, row in df.iterrows():  # Use iterrows to access both columns
        key, value = row['Column1'], row['Column2']
        print(key)
        print(value)

        if start_appending:
            if key == 'AccountHolder':
                break
            account_holder_data[key] = value  # Append the key-value pair
        elif key == 'AccountHolder':
            start_appending = True
            current_key = value

    print(f"Data for AccountHolder {current_key}:")
    for key, value in account_holder_data.items():
        print(f"{key}: {value}")

def transpose_to_new_sheet(source_excel_path, source_sheet_name, output_sheet_name,
                           column_names=['Column1', 'Column2']):
    # Load the DataFrame from the source Excel file and specified sheet
    df = pd.read_excel(source_excel_path, sheet_name=source_sheet_name, usecols=column_names)

    # Transpose the DataFrame and reset the index
    transposed_df = df.T.reset_index(drop=True)
    transposed_df.columns = range(transposed_df.shape[1])

    # Open the workbook
    app = xw.App(visible=False)  # Keep Excel app in background
    workbook = app.books.open(source_excel_path)

    # Check if the output sheet exists and delete it if so
    sheet_names = [sheet.name for sheet in workbook.sheets]
    if output_sheet_name in sheet_names:
        workbook.sheets[output_sheet_name].delete()

    # Add a new sheet for the transposed DataFrame
    sheet = workbook.sheets.add(name=output_sheet_name, after=len(workbook.sheets))

    # Write the transposed DataFrame to the new sheet
    sheet.range('A1').value = transposed_df

    # Save and close the workbook
    workbook.save()
    workbook.close()
    app.quit()  # Ensure the Excel application is properly closed

    print(f"Table transposed and saved to new sheet '{output_sheet_name}' successfully.")

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
                        output_file = os.path.join(output_folder, f"{isin}_{file_base_name}.xlsm")
                        inject_macro('loader', file_base_name, output_file)
                        execute_macro(output_file)
                        transpose_to_new_sheet(output_file, 'Sheet2', 'TRANS', ['Column1', 'Column2'])

                        #print_first_account_holder(output_file)

else:
    print("The specified directory does not exist or is not a directory.")
