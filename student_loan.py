from datetime import date

class StudentLoan:
    def __init__(self, principal, interest):
        self.principal = principal
        self.interest = interest
        self.total = interest+principal

    def printDebtFederal(self):
        print("My federal student loan information: ")
        print("Principal amount: $%.2f" % self.principal)
        print("Interest amount: $%.2f " % self.interest)
        print("Total amount: $%.2f " % self.total)

    def printDebtPrivate(self):
        print("My private student loan information: ")
        print("Principal amount: $%.2f " % self.principal)
        print("Interest amount: $%.2f " % self.interest)
        print("Total amount: $%.2f " % self.total)

def studentLoanSetUp(summarySheet):
    summarySheet['A15'] = "Student Debt Totals"
    summarySheet['A16'] = "Principal"
    summarySheet['A17'] = "Interest"
    summarySheet['A18'] = "Total"
    summarySheet['B16'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B17'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B18'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['B15'] = "Federal"
    summarySheet['C15'] = "Private"
    summarySheet['C16'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['C17'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['C18'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['D15'] = "Total"
    summarySheet['D16'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['D17'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    summarySheet['D18'].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'

def setUpStudentDebtArchive(studentDebtArchive):
    studentDebtArchive['A1'] = "Student Debt Archive"
    studentDebtArchive['B1'] = "Federal Principal"
    studentDebtArchive['C1'] = "Federal Interest"
    studentDebtArchive['D1'] = "Federal Total"
    studentDebtArchive['F1'] = "Private Principal"
    studentDebtArchive['G1'] = "Private Interest"
    studentDebtArchive['H1'] = "Private Total"
    studentDebtArchive['J1'] = "Overall Total"

def allocateSummaryStudentDebt(summarySheet, federal, private):
    summarySheet['B16'] = federal.principal
    summarySheet['B17'] = federal.interest
    summarySheet['B18'] = federal.total
    summarySheet['C16'] = private.principal
    summarySheet['C17'] = private.interest
    summarySheet['C18'] = private.total
    summarySheet['D16'] = federal.principal + private.principal
    summarySheet['D17'] = federal.interest + private.interest
    summarySheet['D18'] = federal.total + private.total

def allocateArchiveStudentDebt(studentDebtArchive, federal, private):
    insertRow = 2
    studentDebtArchive.insert_rows(insertRow)
    studentDebtArchive.cell(row=insertRow, column=1).number_format = 'm/d/yyyy'
    studentDebtArchive.cell(row=insertRow, column=1).value = date.today().strftime("%m/%d/%Y")
    studentDebtArchive.cell(row=insertRow, column=2).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    studentDebtArchive.cell(row=insertRow, column=2).value = federal.principal
    studentDebtArchive.cell(row=insertRow, column=3).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    studentDebtArchive.cell(row=insertRow, column=3).value = federal.interest
    studentDebtArchive.cell(row=insertRow, column=4).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    studentDebtArchive.cell(row=insertRow, column=4).value = federal.total
    studentDebtArchive.cell(row=insertRow, column=6).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    studentDebtArchive.cell(row=insertRow, column=6).value = private.principal
    studentDebtArchive.cell(row=insertRow, column=7).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    studentDebtArchive.cell(row=insertRow, column=7).value = private.interest
    studentDebtArchive.cell(row=insertRow, column=8).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    studentDebtArchive.cell(row=insertRow, column=8).value = private.total
    studentDebtArchive.cell(row=insertRow, column=10).number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
    studentDebtArchive.cell(row=insertRow, column=10).value = federal.total + private.total