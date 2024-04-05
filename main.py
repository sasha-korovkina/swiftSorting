import re
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
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path)
    print(df.head())
    # Initialize an empty list to store the values after 'AccountHolder'
    values_after_account_holder = []

    # Flag to start recording values after we find the first 'AccountHolder'
    start_appending = False

    # Iterate through each value in column 'Column1' (replace 'Column1' with your actual column name if different)
    for value in df['Column1']:  # Access the column directly by its name
        if start_appending:
            values_after_account_holder.append(value)
        if value == 'AccountHolder':
            start_appending = True

    # Remove the first 'AccountHolder' from the list
    if values_after_account_holder and values_after_account_holder[0] == 'AccountHolder':
        values_after_account_holder = values_after_account_holder[1:]

    # Print the collected values
    print(values_after_account_holder)

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
                        # print(f"Excel file saved for ISIN {isin} at: {output_file}")
                        inject_macro(output_file, 'loader', file_base_name)
                        execute_macro(output_file)
                        print_first_account_holder(output_file)

else:
    print("The specified directory does not exist or is not a directory.")
