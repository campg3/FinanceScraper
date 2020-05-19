# Reading an excel file using Python
import bank_account
import investment
import xlrd
from datetime import date
import openpyxl as op
from openpyxl.styles import Alignment
import sys

# sets up the format of the summary sheet
# is only called once, if the file didn't exist when called
def setUpSummary():
    summarySheet['A1'] = "Bank Account Totals"
    summarySheet['A2'] = "Checking"
    summarySheet['A3'] = "Savings"
    summarySheet['A4'] = "Money Market"
    summarySheet['A5'] = "Total"
    summarySheet['B1'] = "Current"
    summarySheet['C1'] = "Delta over 1 month"
    summarySheet['D1'] = "Delta over 3 months"
    summarySheet['E1'] = "Delta over 6 months"
    summarySheet['F1'] = "Delta over 9 months"
    summarySheet['G1'] = "Delta over 1 year"

# set up the archive sheet
def setUpArchive():
    archiveSheet['A1'] = "Bank Archive"
    archiveSheet['B1'] = "Checking"
    archiveSheet['C1'] = "Savings"
    archiveSheet['D1'] = "Money Market"
    archiveSheet['E1'] = "Total"


# enters the account values into the summary sheet and new sheet
def allocateBankAccountValues():
    summarySheet['B2'] = account.checking
    summarySheet['B3'] = account.saving
    summarySheet['B4'] = account.mm
    summarySheet['B5'] = account.sumAccount()
    insertRow = archiveSheet.max_row + 1
    archiveSheet.cell(row=insertRow, column=1).value = date.today().strftime("%m_%d_%Y")
    archiveSheet.cell(row=insertRow, column=2).value = account.checking
    archiveSheet.cell(row=insertRow, column=3).value = account.saving
    archiveSheet.cell(row=insertRow, column=4).value = account.mm
    archiveSheet.cell(row=insertRow, column=5).value = account.sumAccount()

# fills out the summary page
def allocateSummaryBankAccount():
    numEntries = archiveSheet.max_row
    print(numEntries)

    # no analysis available
    if (numEntries-1 == 1):
        summarySheet['C2'] = "N/A"
        summarySheet['C3'] = "N/A"
        summarySheet['C4'] = "N/A"
        summarySheet['C5'] = "N/A"
        summarySheet['D2'] = "N/A"
        summarySheet['D3'] = "N/A"
        summarySheet['D4'] = "N/A"
        summarySheet['D5'] = "N/A"
        summarySheet['E2'] = "N/A"
        summarySheet['E3'] = "N/A"
        summarySheet['E4'] = "N/A"
        summarySheet['E5'] = "N/A"
        summarySheet['F2'] = "N/A"
        summarySheet['F3'] = "N/A"
        summarySheet['F4'] = "N/A"
        summarySheet['F5'] = "N/A"
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    # analysis on 1 month available
    elif (numEntries-1 < 4):
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=numEntries-1, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=numEntries-1, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=numEntries-1, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-1, column=5).value
        summarySheet['D2'] = "N/A"
        summarySheet['D3'] = "N/A"
        summarySheet['D4'] = "N/A"
        summarySheet['D5'] = "N/A"
        summarySheet['E2'] = "N/A"
        summarySheet['E3'] = "N/A"
        summarySheet['E4'] = "N/A"
        summarySheet['E5'] = "N/A"
        summarySheet['F2'] = "N/A"
        summarySheet['F3'] = "N/A"
        summarySheet['F4'] = "N/A"
        summarySheet['F5'] = "N/A"
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    # analysis on 1, 3 month
    elif (numEntries-1 < 7):
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=numEntries-1, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=numEntries-1, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=numEntries-1, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-1, column=5).value
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=numEntries-3, column=2).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=numEntries-3, column=3).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=numEntries-3, column=4).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-3, column=5).value
        summarySheet['E2'] = "N/A"
        summarySheet['E3'] = "N/A"
        summarySheet['E4'] = "N/A"
        summarySheet['E5'] = "N/A"
        summarySheet['F2'] = "N/A"
        summarySheet['F3'] = "N/A"
        summarySheet['F4'] = "N/A"
        summarySheet['F5'] = "N/A"
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    # analysis on 1, 3, 6 month
    elif (numEntries-1 < 10):
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=numEntries-1, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=numEntries-1, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=numEntries-1, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-1, column=5).value
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=numEntries-3, column=2).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=numEntries-3, column=3).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=numEntries-3, column=4).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-3, column=5).value
        summarySheet['E2'] = account.checking - archiveSheet.cell(row=numEntries-6, column=2).value
        summarySheet['E3'] = account.saving - archiveSheet.cell(row=numEntries-6, column=3).value
        summarySheet['E4'] = account.mm - archiveSheet.cell(row=numEntries-6, column=4).value
        summarySheet['E5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-6, column=5).value
        summarySheet['F2'] = "N/A"
        summarySheet['F3'] = "N/A"
        summarySheet['F4'] = "N/A"
        summarySheet['F5'] = "N/A"
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    # analysis on 1, 3, 6, 9 month
    elif (numEntries-1 < 13):
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=numEntries-1, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=numEntries-1, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=numEntries-1, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-1, column=5).value
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=numEntries-3, column=2).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=numEntries-3, column=3).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=numEntries-3, column=4).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-3, column=5).value
        summarySheet['E2'] = account.checking - archiveSheet.cell(row=numEntries-6, column=2).value
        summarySheet['E3'] = account.saving - archiveSheet.cell(row=numEntries-6, column=3).value
        summarySheet['E4'] = account.mm - archiveSheet.cell(row=numEntries-6, column=4).value
        summarySheet['E5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-6, column=5).value
        summarySheet['F2'] = account.checking - archiveSheet.cell(row=numEntries-9, column=2).value
        summarySheet['F3'] = account.saving - archiveSheet.cell(row=numEntries-9, column=3).value
        summarySheet['F4'] = account.mm - archiveSheet.cell(row=numEntries-9, column=4).value
        summarySheet['F5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-9, column=5).value
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    # analysis on 1, 3, 6, 9 month, 1 year
    else:
        print(archiveSheet.cell(row=numEntries-1, column=2).value)
        print(numEntries-1)
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=numEntries-1, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=numEntries-1, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=numEntries-1, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-1, column=5).value
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=numEntries-3, column=2).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=numEntries-3, column=3).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=numEntries-3, column=4).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-3, column=5).value
        summarySheet['E2'] = account.checking - archiveSheet.cell(row=numEntries-6, column=2).value
        summarySheet['E3'] = account.saving - archiveSheet.cell(row=numEntries-6, column=3).value
        summarySheet['E4'] = account.mm - archiveSheet.cell(row=numEntries-6, column=4).value
        summarySheet['E5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-6, column=5).value
        summarySheet['F2'] = account.checking - archiveSheet.cell(row=numEntries-9, column=2).value
        summarySheet['F3'] = account.saving - archiveSheet.cell(row=numEntries-9, column=3).value
        summarySheet['F4'] = account.mm - archiveSheet.cell(row=numEntries-9, column=4).value
        summarySheet['F5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-12, column=5).value
        summarySheet['G2'] = account.checking - archiveSheet.cell(row=numEntries-12, column=2).value
        summarySheet['G3'] = account.saving - archiveSheet.cell(row=numEntries-12, column=3).value
        summarySheet['G4'] = account.mm - archiveSheet.cell(row=numEntries-12, column=4).value
        summarySheet['G5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=numEntries-12, column=5).value


# fit the cells of the sheet parameter ws
def fitCells(ws):
    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                cell.alignment = Alignment(horizontal="center")
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value


# basically the start of "main()"

# argument handling
if (len(sys.argv) == 1):
    print("Usage:\npython finance_analyzer.py <inputFile.xlsx> <summaryFile.xlsx>"), exit()
elif (len(sys.argv) != 3):
    print("Too many arguments.\nUsage:\npython finance_analyzer.py <inputFile.xlsx> <summaryFile.xlsx>"), exit()

# Give the location of the file
loc = (sys.argv[1])

# To open Workbook
try:
    wb = xlrd.open_workbook(loc)
except:
    print("Input file does not exist. Please create the well-formatted input file and use it."), exit()
sheet = wb.sheet_by_index(0)

# Loop through columns
for i in range(sheet.ncols):
    # creates the account information in the BankAccount class
    if (sheet.cell_value(0,i) == "Bank Account"):
        account = bank_account.BankAccount(sheet.cell_value(1, i+1), sheet.cell_value(2, i+1), sheet.cell_value(3, i+1))


# open the excel workbook, if doesn't exist, create it and ensure 'Summary' is first sheet
try:
    writeWB = op.load_workbook(sys.argv[2])
    writeWB.title = sys.argv[2][0:-5]
    summarySheet = writeWB["Summary"]
    archiveSheet = writeWB["Bank Archive"]
except:
    writeWB = op.Workbook()
    writeWB.title = sys.argv[2][0:-5]
    summarySheet = writeWB.create_sheet("Summary")
    archiveSheet = writeWB.create_sheet("Bank Archive")
    if (writeWB.sheetnames[0] != "Summary"):
        writeWB.remove(writeWB[writeWB.sheetnames[0]])
    setUpSummary()
    setUpArchive()



# call the various methods that handle writing to the excel files
allocateBankAccountValues()
allocateSummaryBankAccount()

fitCells(summarySheet)
fitCells(archiveSheet)

# save the file
writeWB.save(writeWB.title + '.xlsx')

account.printAccount()
