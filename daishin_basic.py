import win32com.client

# 대신증권 CybosPlus와 연결이 잘 되었는지 확인
instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
print(instCpCybos.IsConnect)

# 종목 코드 수 반환
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
print(instCpStockCode.GetCount())

# 각 인덱스에 위치하는 종목의 종목명 확인 (첫 번째 인자 값이 0이면 종목 코드, 1이면 종목명, 2이면 FullCode 리턴)
print(instCpStockCode.GetData(1, 0))
print(instCpStockCode.GetData(0, 0))

for i in range(10):
    print(instCpStockCode.GetData(1, i))

# 종목명을 넣어 종목코드, 종목명, 인덱스 출력
stockNum = instCpStockCode.GetCount()

for i in range(stockNum):
    if instCpStockCode.GetData(1, i) == '삼성전자':
        print(instCpStockCode.GetData(0, i))
        print(instCpStockCode.GetData(1, i))
        print(i)

# 종목명을 이용해 종목 코드를 구하는 NameToCode 메소드를 이용
samsungCode = instCpStockCode.NameToCode('삼성전자')
samsungIndex = instCpStockCode.CodeToIndex(samsungCode)
print(samsungCode)
print(samsungIndex)