from pycybos.cputil import *

cpcybos = CpCybos()
is_connect = cpcybos.IsConnect

if is_connect == 0:
    print("CYBOS API 연결 오류")
else:
    print("CYBOS API 연결 완료")


