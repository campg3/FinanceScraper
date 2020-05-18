# Reading an excel file using Python
import bank_account
import loan
import xlrd
from datetime import date
import openpyxl as op
import sys

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
    archiveSheet['A1'] = "Bank Archive"
    archiveSheet['A2'] = "Checking"
    archiveSheet['A3'] = "Savings"
    archiveSheet['A4'] = "Money Market"
    archiveSheet['A5'] = "Total"

def setUpArchiveParam(rowVal):
    archiveSheet.cell(row=rowVal+3, column=1).value = "Checking"
    archiveSheet.cell(row=rowVal+4, column=1).value = "Savings"
    archiveSheet.cell(row=rowVal+5, column=1).value = "Money Market"
    archiveSheet.cell(row=rowVal+6, column=1).value = "Total"

# enters the account values into the summary sheet and new sheet
def allocateBankAccountValues():
    summarySheet['B2'] = account.checking
    summarySheet['B3'] = account.saving
    summarySheet['B4'] = account.mm
    summarySheet['B5'] = account.sumAccount()
    maxRow = archiveSheet.max_row
    maxCol = max((c.column for c in archiveSheet[maxRow] if c.value is not None))
    if (maxCol == 21):
        setUpArchiveParam(maxRow)
        maxRow = maxRow + 6
        maxCol = max((c.column for c in archiveSheet[maxRow] if c.value is not None))
    archiveSheet.cell(row=maxRow-4, column=maxCol+1).value = date.today().strftime("%m_%d_%Y")
    archiveSheet.cell(row=maxRow-3, column=maxCol+1).value = account.checking
    archiveSheet.cell(row=maxRow-2, column=maxCol+1).value = account.saving
    archiveSheet.cell(row=maxRow-1, column=maxCol+1).value = account.mm
    archiveSheet.cell(row=maxRow, column=maxCol+1).value = account.sumAccount()


# fills out the summary page
def allocateSummaryBankAccount():
    maxRow = archiveSheet.max_row
    maxCol = max((c.column for c in archiveSheet[maxRow] if c.value is not None))

    # not enough data for any analysis
    if (maxRow == 5 and maxCol < 3):
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
    # case that allows for analysis for 1 month, 3 month, 6 month, 9 month, 1 year
    elif (maxCol >= 14 or maxRow > 5):
        oneMonthCol = maxCol-1
        oneMonthRow = maxRow
        threeMonthCol = maxCol-3
        threeMonthRow = maxRow
        sixMonthCol = maxCol-6
        sixMonthRow = maxRow
        nineMonthCol = maxCol-9
        nineMonthRow = maxRow
        yearCol = maxCol-12
        yearRow = maxRow
        if (oneMonthCol <= 1):
            oneMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None))
            threeMonthCol = oneMonthCol-2
            sixMonthCol = oneMonthCol-5
            nineMonthCol = oneMonthCol-8
            yearCol = oneMonthCol-11
            oneMonthRow = oneMonthRow-6
            threeMonthRow = oneMonthRow
            sixMonthRow = oneMonthRow
            nineMonthRow = oneMonthRow
            yearRow = oneMonthRow
        elif (threeMonthCol <= 1):
            threeMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-threeMonthCol)
            sixMonthCol = threeMonthCol-3
            nineMonthCol = threeMonthCol-6
            yearCol = threeMonthCol-9
            threeMonthRow = threeMonthRow-6
            sixMonthRow = threeMonthRow
            nineMonthRow = threeMonthRow
            yearRow = threeMonthRow
        elif (sixMonthCol <= 1):
            sixMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-sixMonthCol)
            nineMonthCol = sixMonthCol-3
            yearCol = sixMonthCol-6
            sixMonthRow = maxRow-6
            nineMonthRow = sixMonthRow
            yearRow = sixMonthRow
        elif (nineMonthCol <= 1):
            nineMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-nineMonthCol)
            nineMonthRow = maxRow-6
            yearCol = nineMonthCol-3
            yearRow = nineMonthRow
        elif (yearCol <= 1):
            yearCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-yearCol)
            yearRow = maxRow-6

        summarySheet['D2'] = account.checking - archiveSheet.cell(row=oneMonthRow-3, column=oneMonthCol).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=oneMonthRow-2, column=oneMonthCol).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=oneMonthRow-1, column=oneMonthCol).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=oneMonthRow, column=oneMonthCol).value
        summarySheet['F2'] = account.checking - archiveSheet.cell(row=threeMonthRow-3, column=threeMonthCol).value
        summarySheet['F3'] = account.saving - archiveSheet.cell(row=threeMonthRow-2, column=threeMonthCol).value
        summarySheet['F4'] = account.mm - archiveSheet.cell(row=threeMonthRow-1, column=threeMonthCol).value
        summarySheet['F5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=threeMonthRow, column=threeMonthCol).value
        summarySheet['H2'] = account.checking - archiveSheet.cell(row=sixMonthRow-3, column=sixMonthCol).value
        summarySheet['H3'] = account.saving - archiveSheet.cell(row=sixMonthRow-2, column=sixMonthCol).value
        summarySheet['H4'] = account.mm - archiveSheet.cell(row=sixMonthRow-1, column=sixMonthCol).value
        summarySheet['H5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=sixMonthRow, column=sixMonthCol).value
        summarySheet['J2'] = account.checking - archiveSheet.cell(row=nineMonthRow-3, column=nineMonthCol).value
        summarySheet['J3'] = account.saving - archiveSheet.cell(row=nineMonthRow-2, column=nineMonthCol).value
        summarySheet['J4'] = account.mm - archiveSheet.cell(row=nineMonthRow-1, column=nineMonthCol).value
        summarySheet['J5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=nineMonthRow, column=nineMonthCol).value
        summarySheet['L2'] = account.checking - archiveSheet.cell(row=yearRow-3, column=yearCol).value
        summarySheet['L3'] = account.saving - archiveSheet.cell(row=yearRow-2, column=yearCol).value
        summarySheet['L4'] = account.mm - archiveSheet.cell(row=yearRow-1, column=yearCol).value
        summarySheet['L5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=yearRow, column=yearCol).value
    # case that allows for analysis for 1 month, 3 month, 6 month, 9 month
    elif (maxCol >= 11 or maxRow > 5):
        oneMonthCol = maxCol-1
        oneMonthRow = maxRow
        threeMonthCol = maxCol-3
        threeMonthRow = maxRow
        sixMonthCol = maxCol-6
        sixMonthRow = maxRow
        nineMonthCol = maxCol-9
        nineMonthRow = maxRow
        if (oneMonthCol <= 1):
            oneMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None))
            threeMonthCol = oneMonthCol-2
            sixMonthCol = oneMonthCol-5
            nineMonthCol = oneMonthCol-8
            oneMonthRow = oneMonthRow-6
            threeMonthRow = oneMonthRow
            sixMonthRow = oneMonthRow
            nineMonthRow = oneMonthRow
        elif (threeMonthCol <= 1):
            threeMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-threeMonthCol)
            sixMonthCol = threeMonthCol-3
            nineMonthCol = threeMonthCol-6
            threeMonthRow = threeMonthRow-6
            sixMonthRow = threeMonthRow
            nineMonthRow = threeMonthRow
        elif (sixMonthCol <= 1):
            sixMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-sixMonthCol)
            nineMonthCol = sixMonthCol-3
            sixMonthRow = maxRow-6
            nineMonthRow = sixMonthRow
        elif (nineMonthCol <= 1):
            nineMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-nineMonthCol)
            nineMonthRow = maxRow-6

        summarySheet['D2'] = account.checking - archiveSheet.cell(row=oneMonthRow-3, column=oneMonthCol).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=oneMonthRow-2, column=oneMonthCol).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=oneMonthRow-1, column=oneMonthCol).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=oneMonthRow, column=oneMonthCol).value
        summarySheet['F2'] = account.checking - archiveSheet.cell(row=threeMonthRow-3, column=threeMonthCol).value
        summarySheet['F3'] = account.saving - archiveSheet.cell(row=threeMonthRow-2, column=threeMonthCol).value
        summarySheet['F4'] = account.mm - archiveSheet.cell(row=threeMonthRow-1, column=threeMonthCol).value
        summarySheet['F5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=threeMonthRow, column=threeMonthCol).value
        summarySheet['H2'] = account.checking - archiveSheet.cell(row=sixMonthRow-3, column=sixMonthCol).value
        summarySheet['H3'] = account.saving - archiveSheet.cell(row=sixMonthRow-2, column=sixMonthCol).value
        summarySheet['H4'] = account.mm - archiveSheet.cell(row=sixMonthRow-1, column=sixMonthCol).value
        summarySheet['H5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=sixMonthRow, column=sixMonthCol).value
        summarySheet['J2'] = account.checking - archiveSheet.cell(row=nineMonthRow-3, column=nineMonthCol).value
        summarySheet['J3'] = account.saving - archiveSheet.cell(row=nineMonthRow-2, column=nineMonthCol).value
        summarySheet['J4'] = account.mm - archiveSheet.cell(row=nineMonthRow-1, column=nineMonthCol).value
        summarySheet['J5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=nineMonthRow, column=nineMonthCol).value
        summarySheet['L2'] = "N/A"
        summarySheet['L3'] = "N/A"
        summarySheet['L4'] = "N/A"
        summarySheet['L5'] = "N/A"
    # case that allows for analysis for 1 month, 3 month, 6 month
    elif (maxCol >= 8 or maxRow > 5):
        oneMonthCol = maxCol-1
        oneMonthRow = maxRow
        threeMonthCol = maxCol-3
        threeMonthRow = maxRow
        sixMonthCol = maxCol-6
        sixMonthRow = maxRow
        if (oneMonthCol <= 1):
            oneMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None))
            threeMonthCol = oneMonthCol-2
            sixMonthCol = oneMonthCol-5
            oneMonthRow = oneMonthRow-6
            threeMonthRow = oneMonthRow
            sixMonthRow = oneMonthRow
        elif (threeMonthCol <= 1):
            threeMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-threeMonthCol)
            sixMonthCol = threeMonthCol-3
            threeMonthRow = threeMonthRow-6
            sixMonthRow = threeMonthRow
        elif (sixMonthCol <= 1):
            sixMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-sixMonthCol)
            sixMonthRow = maxRow-6

        summarySheet['D2'] = account.checking - archiveSheet.cell(row=oneMonthRow-3, column=oneMonthCol).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=oneMonthRow-2, column=oneMonthCol).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=oneMonthRow-1, column=oneMonthCol).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=oneMonthRow, column=oneMonthCol).value
        summarySheet['F2'] = account.checking - archiveSheet.cell(row=threeMonthRow-3, column=threeMonthCol).value
        summarySheet['F3'] = account.saving - archiveSheet.cell(row=threeMonthRow-2, column=threeMonthCol).value
        summarySheet['F4'] = account.mm - archiveSheet.cell(row=threeMonthRow-1, column=threeMonthCol).value
        summarySheet['F5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=threeMonthRow, column=threeMonthCol).value
        summarySheet['H2'] = account.checking - archiveSheet.cell(row=sixMonthRow-3, column=sixMonthCol).value
        summarySheet['H3'] = account.saving - archiveSheet.cell(row=sixMonthRow-2, column=sixMonthCol).value
        summarySheet['H4'] = account.mm - archiveSheet.cell(row=sixMonthRow-1, column=sixMonthCol).value
        summarySheet['H5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=sixMonthRow, column=sixMonthCol).value
        summarySheet['J2'] = "N/A"
        summarySheet['J3'] = "N/A"
        summarySheet['J4'] = "N/A"
        summarySheet['J5'] = "N/A"
        summarySheet['L2'] = "N/A"
        summarySheet['L3'] = "N/A"
        summarySheet['L4'] = "N/A"
        summarySheet['L5'] = "N/A"
    # case that allows for analysis for 1 month, 3 month
    elif (maxCol >= 5 or maxRow > 5):
        oneMonthCol = maxCol-1
        oneMonthRow = maxRow
        threeMonthCol = maxCol-3
        threeMonthRow = maxRow
        if (oneMonthCol <= 1):
            oneMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None))
            threeMonthCol = oneMonthCol-2
            oneMonthRow = oneMonthRow-6
            threeMonthRow = oneMonthRow
        elif (threeMonthCol <= 1):
            threeMonthCol = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None)) - (1-threeMonthCol)
            threeMonthRow = threeMonthRow-6
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=oneMonthRow-3, column=oneMonthCol).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=oneMonthRow-2, column=oneMonthCol).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=oneMonthRow-1, column=oneMonthCol).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=oneMonthRow, column=oneMonthCol).value
        summarySheet['F2'] = account.checking - archiveSheet.cell(row=threeMonthRow-3, column=threeMonthCol).value
        summarySheet['F3'] = account.saving - archiveSheet.cell(row=threeMonthRow-2, column=threeMonthCol).value
        summarySheet['F4'] = account.mm - archiveSheet.cell(row=threeMonthRow-1, column=threeMonthCol).value
        summarySheet['F5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=threeMonthRow, column=threeMonthCol).value
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
    elif (maxCol >= 3 or maxRow > 5):
        columnVal = maxCol-1
        if (columnVal <= 1):
            columnVal = max((c.column for c in archiveSheet[maxRow-6] if c.value is not None))
            maxRow = maxRow-6
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=maxRow-3, column=columnVal).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=maxRow-2, column=columnVal).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=maxRow-1, column=columnVal).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=maxRow, column=columnVal).value
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

# save the file
writeWB.save(writeWB.title + '.xlsx')

account.printAccount()
