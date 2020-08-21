from datetime import date

class BankAccount:
    def __init__(self, checking, saving, mm):
        self.checking = checking
        self.saving = saving
        self.mm = mm

    def printAccount(self):
        print("In my bank account, I have:")
        print("Checking: $%.2f" % self.checking)
        print("Savings: $%.2f" % self.saving)
        print("Money Market: $%.2f" % self.mm)
        print("Total: $%.2f" % self.sumAccount())

    def sumAccount(self):
        return self.checking + self.saving + self.mm

# bank account part of the summary set up
def bankSummarySetUp(summarySheet):
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

# enters the account values into the summary sheet and new sheet
# formats to desired format for each inserted cell
def allocateBankAccountValues(summarySheet, archiveSheet, account):
    summarySheet['B2'] = account.checking
    summarySheet['B3'] = account.saving
    summarySheet['B4'] = account.mm
    summarySheet['B5'] = account.sumAccount()
    insertRow = 2
    archiveSheet.insert_rows(insertRow)
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

# fills out the summary page
def allocateSummaryBankAccount(summarySheet, archiveSheet, account):
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
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=3, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=3, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=3, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=3, column=5).value
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
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=3, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=3, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=3, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=3, column=5).value
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=5, column=2).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=5, column=3).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=5, column=4).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=5, column=5).value
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
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=3, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=3, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=3, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=3, column=5).value
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=5, column=2).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=5, column=3).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=5, column=4).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=5, column=5).value
        summarySheet['E2'] = account.checking - archiveSheet.cell(row=8, column=2).value
        summarySheet['E3'] = account.saving - archiveSheet.cell(row=8, column=3).value
        summarySheet['E4'] = account.mm - archiveSheet.cell(row=8, column=4).value
        summarySheet['E5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=8, column=5).value
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
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=3, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=3, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=3, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=3, column=5).value
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=5, column=2).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=5, column=3).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=5, column=4).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=5, column=5).value
        summarySheet['E2'] = account.checking - archiveSheet.cell(row=8, column=2).value
        summarySheet['E3'] = account.saving - archiveSheet.cell(row=8, column=3).value
        summarySheet['E4'] = account.mm - archiveSheet.cell(row=8, column=4).value
        summarySheet['E5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=8, column=5).value
        summarySheet['F2'] = account.checking - archiveSheet.cell(row=11, column=2).value
        summarySheet['F3'] = account.saving - archiveSheet.cell(row=11, column=3).value
        summarySheet['F4'] = account.mm - archiveSheet.cell(row=11, column=4).value
        summarySheet['F5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=11, column=5).value
        summarySheet['G2'] = "N/A"
        summarySheet['G3'] = "N/A"
        summarySheet['G4'] = "N/A"
        summarySheet['G5'] = "N/A"
    # analysis on 1, 3, 6, 9 month, 1 year
    else:
        summarySheet['C2'] = account.checking - archiveSheet.cell(row=3, column=2).value
        summarySheet['C3'] = account.saving - archiveSheet.cell(row=3, column=3).value
        summarySheet['C4'] = account.mm - archiveSheet.cell(row=3, column=4).value
        summarySheet['C5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=3, column=5).value
        summarySheet['D2'] = account.checking - archiveSheet.cell(row=5, column=2).value
        summarySheet['D3'] = account.saving - archiveSheet.cell(row=5, column=3).value
        summarySheet['D4'] = account.mm - archiveSheet.cell(row=5, column=4).value
        summarySheet['D5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=5, column=5).value
        summarySheet['E2'] = account.checking - archiveSheet.cell(row=8, column=2).value
        summarySheet['E3'] = account.saving - archiveSheet.cell(row=8, column=3).value
        summarySheet['E4'] = account.mm - archiveSheet.cell(row=8, column=4).value
        summarySheet['E5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=8, column=5).value
        summarySheet['F2'] = account.checking - archiveSheet.cell(row=11, column=2).value
        summarySheet['F3'] = account.saving - archiveSheet.cell(row=11, column=3).value
        summarySheet['F4'] = account.mm - archiveSheet.cell(row=11, column=4).value
        summarySheet['F5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=11, column=5).value
        summarySheet['G2'] = account.checking - archiveSheet.cell(row=14, column=2).value
        summarySheet['G3'] = account.saving - archiveSheet.cell(row=14, column=3).value
        summarySheet['G4'] = account.mm - archiveSheet.cell(row=14, column=4).value
        summarySheet['G5'] = round(account.sumAccount(), 2) - archiveSheet.cell(row=14, column=5).value

