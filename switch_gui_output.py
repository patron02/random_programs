# This program imports excel data and exports it onto a txt file
# Place the xlsx file in the same path as the py file but not in the python script folder
# You need to have openpyxl library installed

import openpyxl

from openpyxl import Workbook, load_workbook
# Include the data_only=True or else it will read the formulas and not the actual values
wb = load_workbook('data_file.xlsx', data_only=True)
# Selects the sheet name from your excel file
ws = wb["Sheet 1"]

# Open txt file in write mode
with open('switch_output.txt','w') as f:
    # Iterate over all the rows in the sheet starting from row 4
    for row in range(4, ws.max_row + 1):
        
        # Set up variables you want to export to txt
        tag = ws.cell(row=row, column=2).value
        ip = ws.cell(row=row, column=11).value
        l2 = ws.cell(row=row, column=13).value
        port = ws.cell(row=row, column=15).value
        state = ws.cell(row=row, column=17).value
        speed = ws.cell(row=row, column=18).value
        tagged = ws.cell(row=row, column=23).value
        vlan = ws.cell(row=row, column=24).value

        # Print to txt file (change the format if needed)
        f.write(f"Switch: {tag}\n")
        f.write(f"MGMT IP Address: {ip}\n")
        f.write(f"Protocol: {l2}\n")
        f.write(f"Port: {port}\n")
        f.write(f"State: {state}\n")
        f.write(f"Speed: {speed}\n")
        f.write(f"Tagged/Untagged: {tagged}\n")
        f.write(f"VLAN: {vlan}\n")
        f.write("\n")

