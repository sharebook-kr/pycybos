import win32com.client


class CpCybos:
    def __init__(self):
        self.com_obj = win32com.client.Dispatch("cputil.CpCybos")

    @property
    def IsConnect(self):
        return self.com_obj.IsConnect

    @property
    def ServerType(self):
        return self.com_obj.ServerType

    @property
    def LimitRequestRemainTime(self):
        return self.com_obj.LimitRequestRemainTime

    def GetLimitRemainCount(self, limitType):
        value = self.com_obj.GetLimitRemainCount(limitType)
        return value


if __name__ == "__main__":
    cybos = CpCybos()

    # Property
    print(cybos.IsConnect)
    print(cybos.ServerType)
    print(cybos.LimitRequestRemainTime)

    # Method
    print(cybos.GetLimitRemainCount(0))