# Reading an excel file using Python
import bank_account
import investment
import xlrd
from datetime import date
import openpyxl as op
from openpyxl.styles import Alignment
import sys

# bank account part of the summary set up
def bankSummarySetUp():
    summarySheet['A1'] = "Bank Account Totals"
    summarySheet['A2'] = "Checking"
    summarySheet['A3'] = "Savings"
    summarySheet['A4'] = "Money Market"
    summarySheet['A5'] = "Total"
    summarySheet['B1'] = "Current"
    summarySheet['B2'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B3'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B4'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B5'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['C1'] = "Delta over 1 month"
    summarySheet['C2'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['C3'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['C4'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['C5'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['D1'] = "Delta over 3 months"
    summarySheet['D2'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['D3'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['D4'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['D5'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['E1'] = "Delta over 6 months"
    summarySheet['E2'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['E3'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['E4'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['E5'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['F1'] = "Delta over 9 months"
    summarySheet['F2'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['F3'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['F4'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['F5'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['G1'] = "Delta over 1 year"
    summarySheet['G2'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['G3'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['G4'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['G5'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B2'] = "N/A"
    summarySheet['B3'] = "N/A"
    summarySheet['B4'] = "N/A"
    summarySheet['B5'] = "N/A"
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

def investmentSetUp():
    summarySheet['A7'] = "Investment Totals"
    summarySheet['A8'] = "Total Invested"
    summarySheet['A9'] = "Total Profit"
    summarySheet['A10'] = "# of Re-Investments"
    summarySheet['A11'] = "Time since beginning (days)"
    summarySheet['B8'] = 0
    summarySheet['B8'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B9'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B9'] = 0
    summarySheet['B10'] = 0
    summarySheet['B11'] = 0
    summarySheet['C7'] = "Most Recent Investment"
    summarySheet['D8'].number_format = 'm/d/yyyy'
    summarySheet['D9'].number_format = 'm/d/yyyy'
    summarySheet['C8'] = "Date Invested"
    summarySheet['C9'] = "Date To Be Returned"
    summarySheet['D10'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['D13'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['C10'] = "Principal"
    summarySheet['D11'].number_format = '0%'
    summarySheet['C11'] = "Rate"
    summarySheet['C12'] = "Length (days)"
    summarySheet['C13'] = "Profit"
    summarySheet['D8'] = "N/A"
    summarySheet['D9'] = "N/A"
    summarySheet['D10'] = "N/A"
    summarySheet['D11'] = "N/A"
    summarySheet['D12'] = "N/A"
    summarySheet['D13'] = "N/A"

# sets up the format of the summary sheet
# is only called once, if the file didn't exist when called
# formats to desired format for each inserted cell
# defaults initial values to N/A or 0, depending on which section, bankAccount or Invest
def setUpSummary():
    bankSummarySetUp()
    investmentSetUp()


# set up the archive sheet
def setUpArchive():
    archiveSheet['A1'] = "Bank Archive"
    archiveSheet['B1'] = "Checking"
    archiveSheet['C1'] = "Savings"
    archiveSheet['D1'] = "Money Market"
    archiveSheet['E1'] = "Total"

# set up invesment archive sheet
def setUpInvestArchive():
    investArchive['A1'] = "Date Invested"
    investArchive['B1'] = "Date Returned"
    investArchive['C1'] = "Principal"
    investArchive['D1'] = "Rate"
    investArchive['E1'] = "Length of Investment (Days)"
    investArchive['F1'] = "Profit"

# enters the account values into the summary sheet and new sheet
# formats to desired format for each inserted cell
def allocateBankAccountValues():
    summarySheet['B2'] = account.checking
    summarySheet['B3'] = account.saving
    summarySheet['B4'] = account.mm
    summarySheet['B5'] = account.sumAccount()
    insertRow = archiveSheet.max_row + 1 # get the next row, the row we need to insert at
    archiveSheet.cell(row=insertRow, column=1).number_format = 'm/d/yyyy'
    archiveSheet.cell(row=insertRow, column=1).value = date.today().strftime("%m/%d/%Y")
    archiveSheet.cell(row=insertRow, column=2).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    archiveSheet.cell(row=insertRow, column=2).value = account.checking
    archiveSheet.cell(row=insertRow, column=3).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    archiveSheet.cell(row=insertRow, column=3).value = account.saving
    archiveSheet.cell(row=insertRow, column=4).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    archiveSheet.cell(row=insertRow, column=4).value = account.mm
    archiveSheet.cell(row=insertRow, column=5).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    archiveSheet.cell(row=insertRow, column=5).value = account.sumAccount()

# enters invesment values into archive and the current investment values into the Current
# location on the summary sheet
# formats to desired format for each inserted cell
def allocateInvestmentValues():
    summarySheet['D8'] = investment.startDate
    summarySheet['D9'] = investment.endDate
    summarySheet['D10'] = investment.principal
    summarySheet['D11'] = investment.rate
    summarySheet['D12'] = investment.period
    summarySheet['D13'] = round((investment.calculateProfit()), 2)
    insertRow = investArchive.max_row + 1 # get the next row, the row we need to insert at
    investArchive.cell(row=insertRow, column=1).number_format = 'm/d/yyyy'
    investArchive.cell(row=insertRow, column=1).value = investment.startDate
    investArchive.cell(row=insertRow, column=2).number_format = 'm/d/yyyy'
    investArchive.cell(row=insertRow, column=2).value = investment.endDate
    investArchive.cell(row=insertRow, column=3).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    investArchive.cell(row=insertRow, column=3).value = investment.principal
    investArchive.cell(row=insertRow, column=4).number_format = '0%'
    investArchive.cell(row=insertRow, column=4).value = investment.rate
    investArchive.cell(row=insertRow, column=5).value = investment.period
    investArchive.cell(row=insertRow, column=6).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    investArchive.cell(row=insertRow, column=6).value = round((investment.calculateProfit()), 2)

# function that allocates the summary of my investment value to the summary sheet
def allocateSummaryInvestment():
    summarySheet['B8'] = summarySheet['B8'].value + investment.newMoney
    summarySheet['B9'] = round( (summarySheet['B9'].value + (investment.calculateProfit())), 2)
    summarySheet['B10'] = round( (summarySheet['B10'].value + 1), 2)
    summarySheet['B11'] = round( (summarySheet['B11'].value + investment.period), 2)

# fills out the summary page
def allocateSummaryBankAccount():
    numEntries = archiveSheet.max_row

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
        ws.column_dimensions[col].width = value+6


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

bankPresent = False
investmentPresent = False

# Loop through columns
for i in range(sheet.nrows):
    # creates the account information in the BankAccount class
    if (sheet.cell_value(i,0) == "Bank Account"):
        bankPresent = True
        account = bank_account.BankAccount(sheet.cell_value(i+1, 1), sheet.cell_value(i+2, 1), sheet.cell_value(i+3, 1))
    elif (sheet.cell_value(i,0) == "Investment"):
        investmentPresent = True
        investment = investment.Investment(sheet.cell_value(i+1, 1), sheet.cell_value(i+2, 1), sheet.cell_value(i+3, 1),
                                            sheet.cell_value(i+4, 1), sheet.cell_value(i+5, 1), sheet.cell_value(i+6, 1))


# open the excel workbook, if doesn't exist, create it and ensure 'Summary' is first sheet
try:
    writeWB = op.load_workbook(sys.argv[2])
    writeWB.title = sys.argv[2][0:-5]
    summarySheet = writeWB["Summary"]
    archiveSheet = writeWB["Bank Archive"]
    investArchive = writeWB["Investment Archive"]
except:
    writeWB = op.Workbook()
    writeWB.title = sys.argv[2][0:-5] # set the title without the .xlsx
    summarySheet = writeWB.create_sheet("Summary")
    archiveSheet = writeWB.create_sheet("Bank Archive")
    investArchive = writeWB.create_sheet("Investment Archive")
    if (writeWB.sheetnames[0] != "Summary"): # if not summary, delete first
        writeWB.remove(writeWB[writeWB.sheetnames[0]])
    setUpSummary()
    setUpArchive()
    setUpInvestArchive()



# call the various methods that handle writing to the excel files, only if they've
# been included in the inputfile
if (bankPresent == True):
    allocateBankAccountValues()
    allocateSummaryBankAccount()
    account.printAccount()
    print()
if (investmentPresent == True):
    allocateInvestmentValues()
    allocateSummaryInvestment()
    investment.printInformation()
    print()

fitCells(summarySheet)
fitCells(archiveSheet)
fitCells(investArchive)

# save the file
writeWB.save(writeWB.title + '.xlsx')
print("Done")
