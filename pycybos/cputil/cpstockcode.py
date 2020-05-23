import win32com.client


class CpStockCode:
    def __init__(self):
        self.com_obj = win32com.client.Dispatch("cputil.CpStockCode")

    def CodeToName(self, code):
        return self.com_obj.CodeToName(code)

    def NameToCode(self, name):
        code = self.com_obj.NameToCode(name)
        return code

    def CodeToFullCode(self, code):
        ret = self.com_obj.CodeToFullCode(code)
        return ret

    def FullCodeToName(self, fullcode):
        ret = self.com_obj.FullCodeToName(fullcode)
        return ret

    def FullCodeToCode(self, fullcode):
        ret = self.com_obj.FullCodeToCode(fullcode)
        return ret

    def CodeToIndex(self, code):
        ret = self.com_obj.CodeToIndex(code)
        return ret

    def GetCount(self):
        ret = self.com_obj.GetCount()
        return ret

    def GetData(self, type, index):
        ret = self.com_obj.GetData(type, index)
        return ret

    def GetPriceUnit(self, code, base_price, direction_up):
        """
        호가 단위를 얻는 메서드
        :param code: 종목코드
        :param base_price: 기준가격
        :param direction_up: True: 한호가 상승, False: 한호가 하락
        :return:
        """
        ret = self.com_obj.GetPriceUnit(code, base_price, direction_up)
        return ret


if __name__ == "__main__":
    stockcode = CpStockCode()

    # method
    print(stockcode.CodeToName("A005930"))
    print(stockcode.NameToCode("삼성전자"))
    print(stockcode.CodeToFullCode("005930"))
    print(stockcode.FullCodeToName("KR7005930003"))
    print(stockcode.FullCodeToCode("KR7005930003"))
    print(stockcode.CodeToIndex("A005930"))
    print(stockcode.GetCount())
    print(stockcode.GetData(0, 339))
    print(stockcode.GetData(1, 339))
    print(stockcode.GetData(2, 339))
    print(stockcode.GetPriceUnit("005930", 40000, True))
