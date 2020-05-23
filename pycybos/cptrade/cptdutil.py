import win32com.client


class CpTdUtil:
    def __init__(self):
        self.com_obj = win32com.client.Dispatch("cptrade.CpTdUtil")

    @property
    def AccountNumber(self):
        return self.com_obj.AccountNumber

    def TradeInit(self):
        self.com_obj.TradeInit()


if __name__ == "__main__":
    cptduitl = CpTdUtil()
    cptduitl.TradeInit()
    print(cptduitl.AccountNumber)