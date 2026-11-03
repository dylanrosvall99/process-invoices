# Insert your python code here
import pandas as pd
import openpyxl
from analyticskit import *
from datetime import datetime as dt
from pathlib import Path

# Select the invoice data file
while True:
    invoice_data_path = select_file(
        title="Select Invoice-Data.xlsx."
    )

    if invoice_data_path and Path(invoice_data_path).name == "Invoice-Data.xlsx":
        break
    else:
        show_message("Error", "Please select Invoice-Data.xlsx.")


# Select the invoice template file
while True:
    invoice_template_path = select_file(
        title="Select Invoice-Template.xlsx."
    )

    if invoice_template_path and Path(invoice_template_path).name == "Invoice-Template.xlsx":
        break
    else:
        show_message("Error", "Please select Invoice-Template.xlsx.")


# Get the folder where completed invoices will be saved
directory = str(Path(invoice_template_path).parent) + "/"


# Read the invoice data
invoice_data = pd.read_excel(invoice_data_path)

for index,row in invoice_data.iterrows():
    
    invoice_template = openpyxl.load_workbook(invoice_template_path)
    
    selected_sheet = invoice_template['ACCInvoice']
    
    selected_sheet['C3'].value = 'Advanced Accounting Analytics'
    selected_sheet['C4'].value = '1234 Python St., Provo, UT 84602'
    selected_sheet['C5'].value = '(801) 555-0123'
    selected_sheet['C6'].value = 'pythonmaster@byu.edu'
    
    selected_sheet['G6'].value = dt.today().strftime('%m/%d/%Y')
    selected_sheet['G8'].value = row['Invoice Number']
    
    selected_sheet['B14'].value = row['Name']
    selected_sheet['B15'].value = row['Company Name']
    selected_sheet['B16'].value = f"{row['Street Address']}, {row['City']}, {row['State']} {row['Zip Code']}"
    selected_sheet['B17'].value = row['Phone']
    selected_sheet['B18'].value = row['Email Address']
    
    if pd.isna(row['Item1']) == False:
        selected_sheet['B22'].value = row['Item1']
        selected_sheet['E22'].value = row['Quantity1']
        selected_sheet['F22'].value = row['Price1']
        
    if pd.isna(row['Item2']) == False:
        selected_sheet['B23'].value = row['Item2']
        selected_sheet['E23'].value = row['Quantity2']
        selected_sheet['F23'].value = row['Price2']
        
    if pd.isna(row['Item3']) == False:
        selected_sheet['B24'].value = row['Item3']
        selected_sheet['E24'].value = row['Quantity3']
        selected_sheet['F24'].value = row['Price3']
        
    invoice_template.save(f"{directory}Invoice_{row['Invoice Number']}.xlsx")