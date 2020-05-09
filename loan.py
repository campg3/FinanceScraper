class Loan:
    def __init__(self, principal, rate, period):
        self.principal = principal
        self.rate = rate
        self.period = period

    def calculatePayment(self):
        return (((self.principal*self.rate)/365)*self.period)+self.principal


    def printPayment(self):
        print("My loan information:")
        print("Payment: %.2f" % self.calculatePayment())
