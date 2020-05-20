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
        print("My loan information:")
        print("Principal: $%.2f" % self.principal)
        print("Rate: %.2f%%" % (self.rate*100))
        print("Length of Investment (days): %.0f" % self.period)
        print("Profit: $%.2f" % (self.calculatePayment()-self.principal))
