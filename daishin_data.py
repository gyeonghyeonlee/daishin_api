import win32com.client
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")

# GetStockListByMarket : 시장 구분에 따라 주식 종목을 리스트 형태로 제공
codeList = instCpCodeMgr.GetStockListByMarket(1)
print(codeList)

# 각 종목의 종목명을 구하고, 종목 코드를 key로 종목 이름을 value로 사전에 추가
kospi = {}
for code in codeList:
    name = instCpCodeMgr.CodeToName(code)
    kospi[code] = name

f = open('C:\\Users\\gyeon\\Desktop\\kospi.csv', 'w')
for key, value in kospi.items():
    f.write("%s, %s\n" %(key, value))
f.close()

# ETF나 ETN 종목을 제외하고 순수하게 유가증권시장에 상장된 종목에 대한 코드
# 유가증권시장의 종목에 대해 인덱스, 종목 코드, 부 구분 코드, 종목명 출력
for i, code in enumerate(codeList):
    secondCode = instCpCodeMgr.GetStockSectionKind(code)
    name = instCpCodeMgr.CodeToName(code)
    print(i, code, secondCode, name)