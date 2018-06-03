import win32com.client

class CpCodeMgr:
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpUtil.CpCodeMgr")

    def CodeToName(self, code):
        return self.obj.CodeToName(code)

    def GetStockYdOpenPrice(self, code):
        return self.obj.GetStockYdOpenPrice(code)

    def GetStockYdHighPrice(self, code):
        return self.obj.GetStockYdHighPrice(code)

    def GetStockYdLowPrice(self, code):
        return self.obj.GetStockYdLowPrice(code)

    def GetStockYdClosePrice(self, code):
        return self.obj.GetStockYdClosePrice(code)

    def GetStockListByMarket(self, market):
        return self.obj.GetStockListByMarket(market)

if __name__ == "__main__":
    codemgr = CpCodeMgr()
    print(codemgr.GetStockYdOpenPrice("039490"))
    print(codemgr.GetStockYdHighPrice("039490"))
    print(codemgr.GetStockYdLowPrice("039490"))
    print(codemgr.GetStockYdClosePrice("039490"))

    kospi = codemgr.GetStockListByMarket(1)         # CPC_MARKET_KOSPI
    kosdaq= codemgr.GetStockListByMarket(2)         # CPC_MARKET_KOSDAQ
    print(kospi)
    print(kosdaq)
