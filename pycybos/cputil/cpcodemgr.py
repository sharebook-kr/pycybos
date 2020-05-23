import win32com.client


class CpCodeMgr:
    def __init__(self):
        self.com_obj = win32com.client.Dispatch("cputil.CpCodeMgr")

    def CodeToName(self, code):
        return self.com_obj.CodeToName(code)

    def GetStockMarginRate(self, code):
        return self.com_obj.GetStockMarginRate(code)

    def GetStockMemeMin(self, code):
        return self.com_obj.GetStockMemeMin(code)

    def GetStockIndustryCode(self, code):
        return self.com_obj.GetStockIndustryCode(code)

    def GetStockMarketKind(self, code):
        return self.com_obj.GetStockMarketKind(code)

    def GetStockControlKind(self, code):
        return self.com_obj.GetStockControlKind(code)

    def GetStockSupervisionKind(self, code):
        return self.com_obj.GetStockSupervisionKind(code)

    def GetStockStatusKind(self, code):
        return self.com_obj.GetStockStatusKind(code)

    def GetStockCapital(self, code):
        return self.com_obj.GetStockCapital(code)

    def GetStockFiscalMonth(self, code):
        return self.com_obj.GetStockFiscalMonth(code)

    def GetStockGroupCode(self, code):
        return self.com_obj.GetStockGroupCode(code)

    def GetStockKospi200Kind(self, code):
        return self.com_obj.GetStockKospi200Kind(code)

    def GetStockSectionKind(self, code):
        return self.com_obj.GetStockSectionKind(code)

    def GetStockLacKind(self, code):
        return self.com_obj.GetStockLacKind(code)

    def GetStockListedDate(self, code):
        return self.com_obj.GetStockListedDate(code)

    def GetStockMaxPrice(self, code):
        return self.com_obj.GetStockMaxPrice(code)

    def GetStockMinPrice(self, code):
        return self.com_obj.GetStockMinPrice(code)

    def GetStockParPrice(self, code):
        return self.com_obj.GetStockParPrice(code)

    def GetStockStdPrice(self, code):
        return self.com_obj.GetStockStdPrice(code)

    def GetStockYdOpenPrice(self, code):
        return self.com_obj.GetStockYdOpenPrice (code)

    def GetStockYdHighPrice(self, code):
        return self.com_obj.GetStockYdHighPrice (code)

    def GetStockYdLowPrice(self, code):
        return self.com_obj.GetStockYdLowPrice (code)

    def GetStockYdClosePrice(self, code):
        return self.com_obj.GetStockYdClosePrice(code)

    def IsStockCreditEnable(self, code):
        return self.com_obj.IsStockCreditEnable(code)

    def GetStockParPriceChageType (self, code):
        return self.com_obj.GetStockParPriceChageType(code)

    def IsSPAC(self, code):
        return self.com_obj.IsSPAC(code)

    def GetMiniFutureList(self, code):
        return self.com_obj.GetMiniFutureList(code)

    def GetMiniOptionList(self, code):
        return self.com_obj.GetMiniOptionList(code)

    def ReLoadPortData(self):
        return self.com_obj.ReLoadPortData()

    def GetStockElwBasketCodeList(self, code):
        return self.com_obj.GetStockElwBasketCodeList(code)

    def GetStockElwBasketCompList(self, code):
        return self.com_obj.GetStockElwBasketCompList(code)

    def GetStockListByMarket(self, code):
        return self.com_obj.GetStockListByMarket(code)

    def GetGroupCodeList(self, code):
        return self.com_obj.GetGroupCodeList(code)

    def GetGroupName(self, code):
        return self.com_obj.GetGroupName (code)

    def GetIndustryList(self):
        return self.com_obj.GetIndustryList()

    def GetIndustryName(self, code):
        return self.com_obj.GetIndustryName (code)

    def GetMemberList(self):
        return self.com_obj.GetMemberList()

    def GetMemberName(self, code):
        return self.com_obj.GetMemberName(code)

    def GetKosdaqIndustry1List (self):
        return self.com_obj.GetKosdaqIndustry1List ()

    def GetKosdaqIndustry2List(self):
        return self.com_obj.GetKosdaqIndustry2List()

    def GetMarketStartTime(self):
        return self.com_obj.GetMarketStartTime()

    def GetMarketEndTime(self):
        return self.com_obj.GetMarketEndTime()

    def IsFrnMember(self, code):
        return self.com_obj.IsFrnMember(code)

    #--------------------------------------------------------------------------
    # 해외선물
    #--------------------------------------------------------------------------
    def GetTickUnit(self, code):
        return self.com_obj.GetTickUnit(code)

    def GetTickValue(self, code):
        return self.com_obj.GetTickValue(code)

    def OvFutGetAllCodeList(self):
        return self.com_obj.OvFutGetAllCodeList()

    def OvFutGetExchList(self):
        return self.com_obj.OvFutGetExchList()

    def OvFutCodeToName(self, code):
        return self.com_obj.OvFutCodeToName(code)

    def OvFutGetExchCode(self, code):
        return self.com_obj.OvFutGetExchCode(code)

    def OvFutGetLastTradeDate(self, code):
        return self.com_obj.OvFutGetLastTradeDate(code)

    def OvFutGetProdCode(self, code):
        return self.com_obj.OvFutGetProdCode(code)

    def GetStartTime(self, code):
        return self.com_obj.GetStartTime(code)

    def GetEndTime(self, code):
        return self.com_obj.GetEndTime(code)

    def IsTradeCondition(self, code):
        return self.com_obj.IsTradeCondition(code)


if __name__ == "__main__":
    codemgr = CpCodeMgr()

    # method
    print(codemgr.CodeToName("A005930"))
    print(codemgr.GetStockMarginRate("A005930"))
    print(codemgr.GetStockMemeMin("A005930"))
    print(codemgr.GetStockIndustryCode("A005930"))
    print(codemgr.GetStockMarketKind("A005930"))
    print(codemgr.GetStockControlKind("A005930"))
    print(codemgr.GetStockSupervisionKind("A005930"))
    print(codemgr.GetStockStatusKind("A005930"))
    print(codemgr.GetStockCapital("A005930"))
    print(codemgr.GetStockFiscalMonth("A005930"))
    print(codemgr.GetStockGroupCode("A005930"))
    print(codemgr.GetStockKospi200Kind("A005930"))
    print(codemgr.GetStockSectionKind("A005930"))
    print(codemgr.GetStockLacKind("A005930"))
    print(codemgr.GetStockListedDate("A005930"))
    print(codemgr.GetStockMaxPrice("A005930"))
    print(codemgr.GetStockMinPrice("A005930"))
    print(codemgr.GetStockParPrice("A005930"))
    print(codemgr.GetStockStdPrice("A005930"))
    print(codemgr.GetStockYdOpenPrice("A005930"))
    print(codemgr.GetStockYdHighPrice("A005930"))
    print(codemgr.GetStockYdLowPrice("A005930"))
    print(codemgr.GetStockYdClosePrice("A005930"))



