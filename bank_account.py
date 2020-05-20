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
