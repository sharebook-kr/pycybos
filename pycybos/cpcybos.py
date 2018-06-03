import win32com.client

class CpCybos:
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpUtil.CpCybos")

    @property
    def IsConnect(self):
        return self.obj.IsConnect

    @property
    def ServerType(self):
        return self.obj.ServerType

    @property
    def LimitRequestRemainTime(self):
        return self.obj.LimitRequestRemainTime


if __name__ == "__main__":
    cpcybos = CpCybos()
    print(cpcybos.IsConnect)
    print(cpcybos.ServerType)
    print(cpcybos.LimitRequestRemainTime)
