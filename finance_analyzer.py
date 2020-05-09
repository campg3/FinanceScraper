# Reading an excel file using Python
import bank_account
import loan
import xlrd
import xlwt
from datetime import date
from xlutils.copy import copy
import openpyxl as op

def setUpCurrentAndNew():
    newSheet['A1'] = "Bank Account"
    currentSheet['A1'] = "Bank Account"
    newSheet['A2'] = "Checking"
    currentSheet['A2'] = "Checking"
    newSheet['A3'] = "Savings"
    currentSheet['A3'] = "Savings"
    newSheet['A4'] = "Money Market"
    currentSheet['A4'] = "Money Market"
    newSheet['A5'] = "Total"
    currentSheet['A5'] = "Total"

def allocateBankAccountValues():
    newSheet['B2'] = account.checking
    currentSheet['B2'] = account.checking
    newSheet['B3'] = account.saving
    currentSheet['B3'] = account.saving
    newSheet['B4'] = account.mm
    currentSheet['B4'] = account.mm
    newSheet['B5'] = account.sumAccount()
    currentSheet['B5'] = account.sumAccount()

# Give the location of the file
loc = ("test.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
for i in range(sheet.ncols):
    if (sheet.cell_value(0,i) == "Bank Account"):
        account = bank_account.BankAccount(sheet.cell_value(1, i+1), sheet.cell_value(2, i+1), sheet.cell_value(3, i+1))
    if (sheet.cell_value(0,i) == "Loan"):
        loanObject = loan.Loan(sheet.cell_value(1, i+1), sheet.cell_value(2, i+1), sheet.cell_value(3, i+1))
        i += 1

today = date.today()
d1 = today.strftime("%d_%m_%Y")

writeWB = op.load_workbook("FinanceSummary.xlsx")
try:
    possibleDelete = writeWB["Sheet 1"]
    writeWB.remove_sheet(possibleDelete)
except:
    print("No 'Sheet 1', no worries")

newSheet = writeWB.create_sheet(d1)
currentSheet = writeWB["Current"]

setUpCurrentAndNew()
allocateBankAccountValues()

writeWB.save("FinanceSummary.xlsx")

account.printAccount()
loanObject.printPayment()
