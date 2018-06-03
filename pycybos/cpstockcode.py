import win32com.client

class CpStockCode:
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpUtil.CpStockCode")

    def CodeToName(self, code):
        return self.obj.CodeToName(code)

    def NameToCode(self, name):
        return self.obj.NameToCode(name)

    def CodeToFullCode(self, code):
        return self.obj.CodeToFullCode(code)

    def FullCodeToName(self, fullcode):
        return self.obj.FullCodeToName(fullcode)

    def FullCodeToCode(self, fullcode):
        return self.obj.FullCodeToCode(fullcode)

    def CodeToIndex(self, code):
        return self.obj.CodeToIndex(code)

    def GetCount(self):
        return self.obj.GetCount()

    def GetData(self, type, index):
        return self.obj.GetData(type, index)

    def GetPriceUnit(self, code, basePrice, directionUp):
        return self.obj.GetPriceUnit(code, basePrice, directionUp)

if __name__ == "__main__":
    cpstockcode = CpStockCode()
    print(cpstockcode.CodeToName("006800"))
    print(cpstockcode.NameToCode("삼성전자"))
    print(cpstockcode.CodeToFullCode("006800"))
    print(cpstockcode.FullCodeToName("KR7006800007"))
    print(cpstockcode.FullCodeToCode("KR7006800007"))
    print(cpstockcode.CodeToIndex("006800"))
    print(cpstockcode.GetCount())
    print(cpstockcode.GetData(type=0, index=0))
