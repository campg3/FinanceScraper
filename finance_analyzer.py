# Reading an excel file using Python
import bank_account
import loan
import xlrd
import xlwt
from datetime import date
from xlutils.copy import copy
import openpyxl as op

def setUpNew():
    newSheet['A1'] = "Bank Account"
    newSheet['A2'] = "Checking"
    newSheet['A3'] = "Savings"
    newSheet['A4'] = "Money Market"
    newSheet['A5'] = "Total"

def setUpSummary():
    summarySheet['A1'] = "Current Bank Account"
    summarySheet['A2'] = "Checking"
    summarySheet['A3'] = "Savings"
    summarySheet['A4'] = "Money Market"
    summarySheet['A5'] = "Total"
    summarySheet['C1'] = "Delta in 1 month"
    summarySheet['E1'] = "Delta in 6 months"
    summarySheet['G1'] = "Delta in 1 year"

def allocateBankAccountValues():
    newSheet['B2'] = account.checking
    summarySheet['B2'] = account.checking
    newSheet['B3'] = account.saving
    summarySheet['B3'] = account.saving
    newSheet['B4'] = account.mm
    summarySheet['B4'] = account.mm
    newSheet['B5'] = account.sumAccount()
    summarySheet['B5'] = account.sumAccount()

def allocateSummaryBankAccount():
    sheetList = writeWB.sheetnames
    length = len(sheetList)
    if (length <= 2):
        summarySheet['C2'] = "N/A"
        summarySheet['C3'] = "N/A"
        summarySheet['C4'] = "N/A"
        summarySheet['C5'] = "N/A"
        summarySheet['E2'] = "N/A"
        summarySheet['E3'] = "N/A"
        summarySheet['E4'] = "N/A"
        summarySheet['E5'] = "N/A"
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    elif (length <= 6):
        oneMonthAgo = writeWB[sheetList[length-2]]
        summarySheet['C2'] = account.checking - oneMonthAgo['B2'].value
        summarySheet['C3'] = account.saving - oneMonthAgo['B3'].value
        summarySheet['C4'] = account.mm - oneMonthAgo['B4'].value
        summarySheet['C5'] = round(account.sumAccount(), 2) - oneMonthAgo['B5'].value
        summarySheet['E2'] = "N/A"
        summarySheet['E3'] = "N/A"
        summarySheet['E4'] = "N/A"
        summarySheet['E5'] = "N/A"
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    elif (length <= 12):
        oneMonthAgo = writeWB[sheetList[length-2]]
        sixMonthAgo = writeWB[sheetList[length-7]]
        summarySheet['C2'] = account.checking - oneMonthAgo['B2'].value
        summarySheet['C3'] = account.saving - oneMonthAgo['B3'].value
        summarySheet['C4'] = account.mm - oneMonthAgo['B4'].value
        summarySheet['C5'] = round(account.sumAccount(), 2) - oneMonthAgo['B5'].value
        summarySheet['E2'] = account.checking - sixMonthAgo['B2'].value
        summarySheet['E3'] = account.saving - sixMonthAgo['B3'].value
        summarySheet['E4'] = account.mm - sixMonthAgo['B4'].value
        summarySheet['E5'] = round(account.sumAccount(), 2) - sixMonthAgo['B5'].value
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    else:
        oneMonthAgo = writeWB[sheetList[length-2]]
        sixMonthAgo = writeWB[sheetList[length-7]]
        oneYearAgo = writeWB[sheetList[length-13]]
        print(oneYearAgo['B5'].value)
        summarySheet['C2'] = account.checking - oneMonthAgo['B2'].value
        summarySheet['C3'] = account.saving - oneMonthAgo['B3'].value
        summarySheet['C4'] = account.mm - oneMonthAgo['B4'].value
        summarySheet['C5'] = round(account.sumAccount(), 2) - oneMonthAgo['B5'].value
        summarySheet['E2'] = account.checking - sixMonthAgo['B2'].value
        summarySheet['E3'] = account.saving - sixMonthAgo['B3'].value
        summarySheet['E4'] = account.mm - sixMonthAgo['B4'].value
        summarySheet['E5'] = round(account.sumAccount(), 2) - sixMonthAgo['B5'].value
        summarySheet['G2'] = account.checking - oneYearAgo['B2'].value
        summarySheet['G3'] = account.saving - oneYearAgo['B3'].value
        summarySheet['G4'] = account.mm - oneYearAgo['B4'].value
        summarySheet['G5'] = round(account.sumAccount(), 2) - oneYearAgo['B5'].value

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
    writeWB.remove(writeWB["Sheet 1"])
    summarySheet = writeWB.create_sheet("Summary")
    setUpSummary()
except:
    print("No 'Sheet 1', no worries")

newSheet = writeWB.create_sheet(d1)
summarySheet = writeWB["Summary"]

setUpNew()
allocateBankAccountValues()
allocateSummaryBankAccount()

writeWB.save("FinanceSummary.xlsx")

account.printAccount()
loanObject.printPayment()
