# 대량 거래 (거래량이 1000% 이상 급증) 종목
# 대량 거래 시점에서 PBR이 4보다 작은 종목

import win32com.client

# 객체 생성
instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

# SetInputValue
instStockChart.SetInputValue(0, "A005930") # 첫 번째 인자 값은 입력 데이터 타입(0: 종목코드), 두 번째 인자 값은 입력 데이터 값
instStockChart.SetInputValue(1, ord('2')) # 조회할 기간, 기간으로 요청할 때는 '1', 개수료 요청할 때는 '2'
instStockChart.SetInputValue(4, 60) # 요청 개수, 4는 요청 개수라는 타입을 의미, 60이 실제로 요청할 데이터의 개수 (최근 거래일로부터 60일치에 해당하는 데이터)
instStockChart.SetInputValue(5, 8) # 요청할 데이터의 종류, 종가에 해당하는 값은 5
instStockChart.SetInputValue(6, ord('D')) # 데이터 차트의 종류, 일 단위의 데이터를 가져오기 위해 'D' 입력
instStockChart.SetInputValue(9, ord('1')) # 수정 주가의 반영 여부

# BlockRequest
instStockChart.BlockRequest()

# GetData
volumes = []
numData = instStockChart.GetHeaderValue(3)

for i in range(numData):
    volume = instStockChart.GetDataValue(0, i)
    volumes.append(volume)
print(volumes)

# 평균 거래량 계산
averageVolume = (sum(volumes) - volumes[0]) / (len(volumes) -1)

if (volumes[0] > averageVolume * 10):
    print("대박주")
else:
    print("일반주", volumes[0] / averageVolume)


##### 유가증권시장의 전 종목 중 거래량이 1000% 이상 급증한 종목 찾기

def CheckVolume(instStockChart, code):
    # SetInputValue
    instStockChart.SetInputValue(0, code)  # 첫 번째 인자 값은 입력 데이터 타입(0: 종목코드), 두 번째 인자 값은 입력 데이터 값
    instStockChart.SetInputValue(1, ord('2'))  # 조회할 기간, 기간으로 요청할 때는 '1', 개수료 요청할 때는 '2'
    instStockChart.SetInputValue(4, 60)  # 요청 개수, 4는 요청 개수라는 타입을 의미, 60이 실제로 요청할 데이터의 개수 (최근 거래일로부터 60일치에 해당하는 데이터)
    instStockChart.SetInputValue(5, 8)  # 요청할 데이터의 종류, 종가에 해당하는 값은 5
    instStockChart.SetInputValue(6, ord('D'))  # 데이터 차트의 종류, 일 단위의 데이터를 가져오기 위해 'D' 입력
    instStockChart.SetInputValue(9, ord('1'))  # 수정 주가의 반영 여부

    # BlockRequest
    instStockChart.BlockRequest()

    # GetData
    volumes = []
    numData = instStockChart.GetHeaderValue(3)

    for i in range(numData):
        volume = instStockChart.GetDataValue(0, i)
        volumes.append(volume)

    # 평균 거래량 계산
    averageVolume = (sum(volumes) - volumes[0]) / (len(volumes) - 1)

    if (volumes[0] > averageVolume * 10):
        return 1
    else:
        return 0

if __name__ == "__main__":
    instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = instCpCodeMgr.GetStockListByMarket(1)
    buyList = []
    for code in codeList:
        if CheckVolume(instStockChart, code)==1:
            buyList.append(code)
            print(code)