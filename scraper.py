import requests
import openpyxl
from datetime import datetime

# Open workbook
wb = openpyxl.load_workbook("Insurance_Hiring_Dashboard.xlsx")

today = datetime.now().strftime("%Y-%m-%d")
sheet_name = f"HK Jobs - {today}"

if sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
else:
    ws = wb.create_sheet(sheet_name)

    ws["A1"] = "Company"
    ws["B1"] = "Job Title"
    ws["C1"] = "Posted"
    ws["D1"] = "Link"

row = ws.max_row + 1

# Manulife
url = "https://manulife.wd3.myworkdayjobs.com/wday/cxs/manulife/MFCJH_Jobs/jobs"

payload = {
    "appliedFacets": {
        "locations": [
            "b9c75e3090f74554b537eda632d6eb33",
            "90905028607c01f220226a2382575ce9"
        ]
    },
    "limit": 100,
    "offset": 0,
    "searchText": ""
}

headers = {
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0"
}

try:
    response = requests.post(
        url,
        json=payload,
        headers=headers,
        timeout=30
    )

    data = response.json()

    for job in data.get("jobPostings", []):

        title = job.get("title", "")
        posted = job.get("postedOn", "")
        link = job.get("externalPath", "")

        ws.cell(row, 1, "Manulife")
        ws.cell(row, 2, title)
        ws.cell(row, 3, posted)
        ws.cell(row, 4, link)

        row += 1

except Exception as e:

    ws.cell(row, 1, "ERROR")
    ws.cell(row, 2, str(e))

wb.save("Insurance_Hiring_Dashboard.xlsx")

print("Manulife jobs added")
