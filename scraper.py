import openpyxl
from datetime import datetime

print("Insurance Hiring Dashboard")

today = datetime.now().strftime("%Y-%m-%d")

wb = openpyxl.load_workbook("Insurance_Hiring_Dashboard.xlsx")

sheet_name = f"HK Jobs - {today}"

if sheet_name not in wb.sheetnames:
    ws = wb.create_sheet(sheet_name)

    ws["A1"] = "Company"
    ws["B1"] = "Job Title"
    ws["C1"] = "Posted"
    ws["D1"] = "Link"

wb.save("Insurance_Hiring_Dashboard.xlsx")

print(f"Created sheet: {sheet_name}")
