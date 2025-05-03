import openpyxl 
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Score-Tracker-Data"

ws.append(["Name", "Grade", "Remarks"])
ws.geometry("1366x768")
wb.save("gradinglist.xlsx")
print("Excel file created successfully.")