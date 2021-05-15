import win32com.client

# 과거 데이터를 조회하는데 필요한 StockChart 클래스
instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

# 최근 10일간 종가 데이터 구하기
instStockChart.SetInputValue(0, "A005930") # 첫 번째 인자 값은 입력 데이터 타입(0: 종목코드), 두 번째 인자 값은 입력 데이터 값
instStockChart.SetInputValue(1, ord('2')) # 조회할 기간, 기간으로 요청할 때는 '1', 개수료 요청할 때는 '2'
instStockChart.SetInputValue(4, 10) # 요청 개수, 4는 요청 개수라는 타입을 의미, 10이 실제로 요청할 데이터의 개수 (최근 거래일로부터 10일치에 해당하는 데이터)
#instStockChart.SetInputValue(5, 5) # 요청할 데이터의 종류, 종가에 해당하는 값은 5
instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 8)) # 시가, 고가, 저가, 종가, 거래량
instStockChart.SetInputValue(6, ord('D')) # 데이터 차트의 종류, 일 단위의 데이터를 가져오기 위해 'D' 입력
instStockChart.SetInputValue(9, ord('1')) # 수정 주가의 반영 여부

instStockChart.BlockRequest() # 데이터 처리 요청

# GetHeaderValue 메소드를 토해 수신한 데이터의 개수 확인
numData = instStockChart.GetHeaderValue(3)
numField = instStockChart.GetHeaderValue(1)

print("일자      시가   고가   저가   종가   거래량")
for i in range(numData):
    for j in range(numField):
        print(instStockChart.GetDataValue(j, i), end=" ")
    print("")