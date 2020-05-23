from pycybos.cputil import *

cpcodemgr = CpCodeMgr()

# 유가증권시장
codes = cpcodemgr.GetStockListByMarket(1)
print(codes)
print(len(codes))

# 코스닥시장
kosdaq_codes = cpcodemgr.GetStockListByMarket(2)
print(kosdaq_codes)
print(len(kosdaq_codes))

# 종목코드로 종목명 얻기
name = cpcodemgr.CodeToName("005930")
print(name)
