import win32com.client
import pandas as pd
import calendar


class StockChart:
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpSysDib.StockChart")

    def get_chart(self, code, chart_type, count=None, start=None, end=None):
        self.obj.SetInputValue(0, code)

        # 요청구분
        if count is None and start is not None and end is not None:
            self.obj.SetInputValue(1, ord('1'))
            self.obj.SetInputValue(2, end)
            self.obj.SetInputValue(3, start)
        else:
            self.obj.SetInputValue(1, ord('2'))
            self.obj.SetInputValue(4, count)

        self.obj.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8, 9, 12, 13, 17, 18, 19, 20, 21])
        self.obj.SetInputValue(6, ord(chart_type))
        self.obj.SetInputValue(9, '1')

        # BlockRequest
        self.obj.BlockRequest()

        cols = self.obj.GetHeaderValue(1)
        columns = self.obj.GetHeaderValue(2)
        rows = self.obj.GetHeaderValue(3)

        data = []
        index = []

        for row in range(rows):
            dic = {}

            # 날짜/시간
            date = str(self.obj.GetDataValue(0, row))
            time = str(self.obj.GetDataValue(1, row))
            if len(time) == 3:
                time = '0' + time

            # 차트 타입에 따라 날짜/시간 처리
            if chart_type == "D":
                index.append(pd.to_datetime(date))
            elif chart_type == "W":
                year = date[0:4]
                month = date[4:6]
                sunday_idx = int(date[6])-1
                number_of_days = calendar.monthrange(int(year), int(month))[1]
                start = year + month + "01"
                end = year + month + str(number_of_days)
                sundays = pd.date_range(start=start, end=end, freq='W-SUN')
                index.append(pd.to_datetime(sundays[sunday_idx]))
            elif chart_type == 'M':
                year = date[0:4]
                month = date[4:6]
                number_of_days = calendar.monthrange(int(year), int(month))[1]
                end = year + month + str(number_of_days)
                index.append(pd.to_datetime(end))
            else:
                index.append(pd.to_datetime(date + time))

            for col in range(2, cols):
                label = columns[col]
                val = self.obj.GetDataValue(col, row)
                dic[label] = val

            data.append(dic)

        df = pd.DataFrame(data, columns=columns[2:], index=index)
        return df[::-1]


if __name__ == "__main__":
    stockchart = StockChart()

    # 개수로 요청
    #print(stockchart.get_chart("A000020", chart_type='D', count=10))       # 일
    #print(stockchart.get_chart("A000020", chart_type='W', count=10))       # 주
    #print(stockchart.get_chart("A000020", chart_type='M', count=10))       # 월
    #print(stockchart.get_chart("A000020", chart_type='m', count=10))       # 분
    #print(stockchart.get_chart("A000020", chart_type='T', count=10))       # 틱

    # 기간으로 요청
    #print(stockchart.get_chart("A000020", chart_type='D', start="20180101", end="20180601"))
    print(stockchart.get_chart("A000020", chart_type='W', start="20180101", end="20180601"))
    #print(stockchart.get_chart("A000020", chart_type='M', start="20180101", end="20180601"))
    #print(stockchart.get_chart("A000020", chart_type='m', start="20180101", end="20180601"))
    #print(stockchart.get_chart("A000020", chart_type='T', start="20180101", end="20180601"))








