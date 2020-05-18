# Reading an excel file using Python
import bank_account
import loan
import xlrd
from datetime import date
import openpyxl as op
import sys

# sets up the format of the new sheet
def setUpNew():
    newSheet['A1'] = "Bank Account"
    newSheet['A2'] = "Checking"
    newSheet['A3'] = "Savings"
    newSheet['A4'] = "Money Market"
    newSheet['A5'] = "Total"

# sets up the format of the summary sheet
# is only called once, if the file didn't exist when called
def setUpSummary():
    summarySheet['A1'] = "Current Bank Account"
    summarySheet['A2'] = "Checking"
    summarySheet['A3'] = "Savings"
    summarySheet['A4'] = "Money Market"
    summarySheet['A5'] = "Total"
    summarySheet['D1'] = "Delta in 1 month"
    summarySheet['F1'] = "Delta in 3 months"
    summarySheet['H1'] = "Delta in 6 months"
    summarySheet['J1'] = "Delta in 9 months"
    summarySheet['L1'] = "Delta in 1 year"

# set up the archive sheet
def setUpArchive():
    archiveSheet['A1'] = "Archive"
    archiveSheet['A2'] = "Checking"
    archiveSheet['A3'] = "Savings"
    archiveSheet['A4'] = "Money Market"
    archiveSheet['A5'] = "Total"

def setUpArchive(rowVal):
    archiveSheet.cell(row=rowVal+3, column=1).value = "Checking"
    archiveSheet.cell(row=rowVal+4, column=1).value = "Savings"
    archiveSheet.cell(row=rowVal+5, column=1).value = "Money Market"
    archiveSheet.cell(row=rowVal+6, column=1).value = "Total"

# enters the account values into the summary sheet and new sheet
def allocateBankAccountValues():
    newSheet['B2'] = account.checking
    summarySheet['B2'] = account.checking
    newSheet['B3'] = account.saving
    summarySheet['B3'] = account.saving
    newSheet['B4'] = account.mm
    summarySheet['B4'] = account.mm
    newSheet['B5'] = account.sumAccount()
    summarySheet['B5'] = account.sumAccount()

# fills out the summary page
def allocateSummaryBankAccount():
    sheetList = writeWB.sheetnames
    sheetList.remove('Summary')
    sheetList.remove('Archive')
    length = len(sheetList)

    # not enough data for any analysis
    if (length < 2):
        summarySheet['D2'] = "N/A"
        summarySheet['D3'] = "N/A"
        summarySheet['D4'] = "N/A"
        summarySheet['D5'] = "N/A"
        summarySheet['F2'] = "N/A"
        summarySheet['F3'] = "N/A"
        summarySheet['F4'] = "N/A"
        summarySheet['F5'] = "N/A"
        summarySheet['H2'] = "N/A"
        summarySheet['H3'] = "N/A"
        summarySheet['H4'] = "N/A"
        summarySheet['H5'] = "N/A"
        summarySheet['J2'] = "N/A"
        summarySheet['J3'] = "N/A"
        summarySheet['J4'] = "N/A"
        summarySheet['J5'] = "N/A"
        summarySheet['L2'] = "N/A"
        summarySheet['L3'] = "N/A"
        summarySheet['L4'] = "N/A"
        summarySheet['L5'] = "N/A"
    # case that allows for analysis for 1 month ago
    elif (length < 4):
        oneMonthAgo = writeWB[sheetList[length-2]]
        summarySheet['D2'] = account.checking - oneMonthAgo['B2'].value
        summarySheet['D3'] = account.saving - oneMonthAgo['B3'].value
        summarySheet['D4'] = account.mm - oneMonthAgo['B4'].value
        summarySheet['D5'] = round(account.sumAccount(), 2) - oneMonthAgo['B5'].value
        summarySheet['F2'] = "N/A"
        summarySheet['F3'] = "N/A"
        summarySheet['F4'] = "N/A"
        summarySheet['F5'] = "N/A"
        summarySheet['H2'] = "N/A"
        summarySheet['H3'] = "N/A"
        summarySheet['H4'] = "N/A"
        summarySheet['H5'] = "N/A"
        summarySheet['J2'] = "N/A"
        summarySheet['J3'] = "N/A"
        summarySheet['J4'] = "N/A"
        summarySheet['J5'] = "N/A"
        summarySheet['L2'] = "N/A"
        summarySheet['L3'] = "N/A"
        summarySheet['L4'] = "N/A"
        summarySheet['L5'] = "N/A"
    # case that allows for analysis for 1 month, 3 month
    elif (length < 7):
        oneMonthAgo = writeWB[sheetList[length-2]]
        threeMonthAgo = writeWB[sheetList[length-4]]
        summarySheet['D2'] = account.checking - oneMonthAgo['B2'].value
        summarySheet['D3'] = account.saving - oneMonthAgo['B3'].value
        summarySheet['D4'] = account.mm - oneMonthAgo['B4'].value
        summarySheet['D5'] = round(account.sumAccount(), 2) - oneMonthAgo['B5'].value
        summarySheet['F2'] = account.checking - threeMonthAgo['B2'].value
        summarySheet['F3'] = account.saving - threeMonthAgo['B3'].value
        summarySheet['F4'] = account.mm - threeMonthAgo['B4'].value
        summarySheet['F5'] = round(account.sumAccount(), 2) - threeMonthAgo['B5'].value
        summarySheet['H2'] = "N/A"
        summarySheet['H3'] = "N/A"
        summarySheet['H4'] = "N/A"
        summarySheet['H5'] = "N/A"
        summarySheet['J2'] = "N/A"
        summarySheet['J3'] = "N/A"
        summarySheet['J4'] = "N/A"
        summarySheet['J5'] = "N/A"
        summarySheet['L2'] = "N/A"
        summarySheet['L3'] = "N/A"
        summarySheet['L4'] = "N/A"
        summarySheet['L5'] = "N/A"
    # case that allows for analysis for 1 month, 3 month, 6 month
    elif (length < 10):
        oneMonthAgo = writeWB[sheetList[length-2]]
        threeMonthAgo = writeWB[sheetList[length-4]]
        sixMonthAgo = writeWB[sheetList[length-7]]
        summarySheet['D2'] = account.checking - oneMonthAgo['B2'].value
        summarySheet['D3'] = account.saving - oneMonthAgo['B3'].value
        summarySheet['D4'] = account.mm - oneMonthAgo['B4'].value
        summarySheet['D5'] = round(account.sumAccount(), 2) - oneMonthAgo['B5'].value
        summarySheet['F2'] = account.checking - threeMonthAgo['B2'].value
        summarySheet['F3'] = account.saving - threeMonthAgo['B3'].value
        summarySheet['F4'] = account.mm - threeMonthAgo['B4'].value
        summarySheet['F5'] = round(account.sumAccount(), 2) - threeMonthAgo['B5'].value
        summarySheet['H2'] = account.checking - sixMonthAgo['B2'].value
        summarySheet['H3'] = account.saving - sixMonthAgo['B3'].value
        summarySheet['H4'] = account.mm - sixMonthAgo['B4'].value
        summarySheet['H5'] = round(account.sumAccount(), 2) - sixMonthAgo['B5'].value
        summarySheet['J2'] = "N/A"
        summarySheet['J3'] = "N/A"
        summarySheet['J4'] = "N/A"
        summarySheet['J5'] = "N/A"
        summarySheet['L2'] = "N/A"
        summarySheet['L3'] = "N/A"
        summarySheet['L4'] = "N/A"
        summarySheet['L5'] = "N/A"
    # case that allows for analysis for 1 month, 3 month, 6 month, 9 month
    elif (length < 13):
        oneMonthAgo = writeWB[sheetList[length-2]]
        threeMonthAgo = writeWB[sheetList[length-4]]
        sixMonthAgo = writeWB[sheetList[length-7]]
        nineMonthAgo = writeWB[sheetList[length-10]]
        summarySheet['D2'] = account.checking - oneMonthAgo['B2'].value
        summarySheet['D3'] = account.saving - oneMonthAgo['B3'].value
        summarySheet['D4'] = account.mm - oneMonthAgo['B4'].value
        summarySheet['D5'] = round(account.sumAccount(), 2) - oneMonthAgo['B5'].value
        summarySheet['F2'] = account.checking - threeMonthAgo['B2'].value
        summarySheet['F3'] = account.saving - threeMonthAgo['B3'].value
        summarySheet['F4'] = account.mm - threeMonthAgo['B4'].value
        summarySheet['F5'] = round(account.sumAccount(), 2) - threeMonthAgo['B5'].value
        summarySheet['H2'] = account.checking - sixMonthAgo['B2'].value
        summarySheet['H3'] = account.saving - sixMonthAgo['B3'].value
        summarySheet['H4'] = account.mm - sixMonthAgo['B4'].value
        summarySheet['H5'] = round(account.sumAccount(), 2) - sixMonthAgo['B5'].value
        summarySheet['J2'] = account.checking - nineMonthAgo['B2'].value
        summarySheet['J3'] = account.saving - nineMonthAgo['B3'].value
        summarySheet['J4'] = account.mm - nineMonthAgo['B4'].value
        summarySheet['J5'] = round(account.sumAccount(), 2) - nineMonthAgo['B5'].value
        summarySheet['L2'] = "N/A"
        summarySheet['L3'] = "N/A"
        summarySheet['L4'] = "N/A"
        summarySheet['L5'] = "N/A"
    # case that allows for analysis for 1 month, 3 month, 6 month, 9 month, 1 year
    else:
        oneMonthAgo = writeWB[sheetList[length-2]]
        threeMonthAgo = writeWB[sheetList[length-4]]
        sixMonthAgo = writeWB[sheetList[length-7]]
        nineMonthAgo = writeWB[sheetList[length-10]]
        oneYearAgo = writeWB[sheetList[length-13]]
        summarySheet['D2'] = account.checking - oneMonthAgo['B2'].value
        summarySheet['D3'] = account.saving - oneMonthAgo['B3'].value
        summarySheet['D4'] = account.mm - oneMonthAgo['B4'].value
        summarySheet['D5'] = round(account.sumAccount(), 2) - oneMonthAgo['B5'].value
        summarySheet['F2'] = account.checking - threeMonthAgo['B2'].value
        summarySheet['F3'] = account.saving - threeMonthAgo['B3'].value
        summarySheet['F4'] = account.mm - threeMonthAgo['B4'].value
        summarySheet['F5'] = round(account.sumAccount(), 2) - threeMonthAgo['B5'].value
        summarySheet['H2'] = account.checking - sixMonthAgo['B2'].value
        summarySheet['H3'] = account.saving - sixMonthAgo['B3'].value
        summarySheet['H4'] = account.mm - sixMonthAgo['B4'].value
        summarySheet['H5'] = round(account.sumAccount(), 2) - sixMonthAgo['B5'].value
        summarySheet['J2'] = account.checking - nineMonthAgo['B2'].value
        summarySheet['J3'] = account.saving - nineMonthAgo['B3'].value
        summarySheet['J4'] = account.mm - nineMonthAgo['B4'].value
        summarySheet['J5'] = round(account.sumAccount(), 2) - nineMonthAgo['B5'].value
        summarySheet['L2'] = account.checking - oneYearAgo['B2'].value
        summarySheet['L3'] = account.saving - oneYearAgo['B3'].value
        summarySheet['L4'] = account.mm - oneYearAgo['B4'].value
        summarySheet['L5'] = round(account.sumAccount(), 2) - oneYearAgo['B5'].value
        # if there are more than 13, the number the allows for 1 year analysis, delete and archive
        if (length > 13):
            maxRow = archiveSheet.max_row
            maxCol = max((c.column for c in archiveSheet[maxRow] if c.value is not None))
            if (maxCol == 21):
                setUpArchive(maxRow)
                maxRow = maxRow + 6
                maxCol = max((c.column for c in archiveSheet[maxRow] if c.value is not None))
            archiveSheet.cell(row=maxRow-4, column=maxCol+1).value = date.today().strftime("%d_%m_%Y")
            archiveSheet.cell(row=maxRow-3, column=maxCol+1).value = writeWB[sheetList[0]]['B2'].value
            archiveSheet.cell(row=maxRow-2, column=maxCol+1).value = writeWB[sheetList[0]]['B3'].value
            archiveSheet.cell(row=maxRow-1, column=maxCol+1).value = writeWB[sheetList[0]]['B4'].value
            archiveSheet.cell(row=maxRow, column=maxCol+1).value = writeWB[sheetList[0]]['B5'].value
            writeWB.remove(writeWB[sheetList[0]])


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

# grab today's date and format it
today = date.today()
d1 = today.strftime("%d_%m_%Y")

# open the excel workbook, if doesn't exist, create it and ensure 'Summary' is first sheet
try:
    writeWB = op.load_workbook(sys.argv[2])
    writeWB.title = sys.argv[2][0:-5]
    summarySheet = writeWB["Summary"]
    archiveSheet = writeWB["Archive"]
except:
    writeWB = op.Workbook()
    writeWB.title = sys.argv[2][0:-5]
    summarySheet = writeWB.create_sheet("Summary")
    archiveSheet = writeWB.create_sheet("Archive")
    if (writeWB.sheetnames[0] != "Summary"):
        writeWB.remove(writeWB[writeWB.sheetnames[0]])
    setUpSummary()
    setUpArchive()

# create a new sheet with title of the current date
newSheet = writeWB.create_sheet(d1)

# call the various methods that handle writing to the excel files
setUpNew()
allocateBankAccountValues()
allocateSummaryBankAccount()

# save the file
writeWB.save(writeWB.title + '.xlsx')

account.printAccount()
