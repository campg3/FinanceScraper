class Investment:
    def __init__(self, startDate, endDate, principal, rate, period, newMoney):
        self.principal = principal
        self.rate = rate
        self.period = period
        self.startDate = startDate
        self.endDate = endDate
        self.newMoney = newMoney

    def calculatePayment(self):
        return (((self.principal*self.rate)/365)*self.period)+self.principal

    def calculateProfit(self):
        return (((self.principal*self.rate)/365)*self.period)

    def printInformation(self):
        print("My investment information:")
        print("Principal: $%.2f" % self.principal)
        print("Rate: %.2f%%" % (self.rate*100))
        print("Length of Investment (days): %.0f" % self.period)
        print("Profit: $%.2f" % (self.calculatePayment()-self.principal))

def investmentSetUp(summarySheet):
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

# set up invesment archive sheet
def setUpInvestArchive(investArchive):
    investArchive['A1'] = "Date Invested"
    investArchive['B1'] = "Date Returned"
    investArchive['C1'] = "Principal"
    investArchive['D1'] = "Rate"
    investArchive['E1'] = "Length of Investment (Days)"
    investArchive['F1'] = "Profit"

# enters invesment values into archive and the current investment values into the Current
# location on the summary sheet
# formats to desired format for each inserted cell
def allocateInvestmentValues(investArchive, investment):
    insertRow = 2
    investArchive.insert_rows(insertRow)
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
def allocateSummaryInvestment(summarySheet, investment):
    summarySheet['B8'] = summarySheet['B8'].value + investment.newMoney
    summarySheet['B9'] = round( (summarySheet['B9'].value + (investment.calculateProfit())), 2)
    summarySheet['B10'] = round( (summarySheet['B10'].value + 1), 2)
    summarySheet['B11'] = round( (summarySheet['B11'].value + investment.period), 2)
    summarySheet['D8'] = investment.startDate
    summarySheet['D9'] = investment.endDate
    summarySheet['D10'] = investment.principal
    summarySheet['D11'] = investment.rate
    summarySheet['D12'] = investment.period
    summarySheet['D13'] = round((investment.calculateProfit()), 2)
