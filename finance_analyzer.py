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
    summarySheet['C1'] = "Delta in 1 month"
    summarySheet['E1'] = "Delta in 6 months"
    summarySheet['G1'] = "Delta in 1 year"

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
    length = len(sheetList)

    # not enough data for any analysis
    if (length < 2):
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
    # case that allows for analysis for 1 month ago
    elif (length < 6):
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
    # case that allows for analysis for 1 month, 6 month
    elif (length < 12):
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
    # case that allows for analysis for 1 month, 6 month, 1 year
    else:
        oneMonthAgo = writeWB[sheetList[length-2]]
        sixMonthAgo = writeWB[sheetList[length-7]]
        oneYearAgo = writeWB[sheetList[length-13]]
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
    summarySheet = writeWB["Summary"]
except:
    writeWB = op.Workbook()
    writeWB.title = sys.argv[2][0:-5]
    summarySheet = writeWB.create_sheet("Summary")
    if (writeWB.sheetnames[0] != "Summary"):
        writeWB.remove(writeWB[writeWB.sheetnames[0]])
    setUpSummary()

# create a new sheet with title of the current date
newSheet = writeWB.create_sheet(d1)

# call the various methods that handle writing to the excel files
setUpNew()
allocateBankAccountValues()
allocateSummaryBankAccount()

# save the file
writeWB.save("FinanceSummary.xlsx")

account.printAccount()
