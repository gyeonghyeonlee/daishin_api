import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
# 업종별 코드 리스트
industryCodeList = instCpCodeMgr.GetIndustryList()

# 각 업종별 코드에 해당하는 업종명 출력
for industryCode in industryCodeList:
    print(industryCode, instCpCodeMgr.GetIndustryName(industryCode))

# 유가증권시장에서 음식료품 업종에 속하는 종목의 종목 코드 리스트를 구한 후 해당 리스트에 속하는 종목의 종목명 출력
targetCodeList = instCpCodeMgr.GetGroupCodeList(5)

for code in targetCodeList:
    print(code, instCpCodeMgr.CodeToName(code))

# 음식료품 업종의 평균 PER 계산
instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

# Get PER
instMarketEye.SetInputValue(0, 67)
instMarketEye.SetInputValue(1, targetCodeList)

# BlockRequest
instMarketEye.BlockRequest()

# GetHeaderValue
numStock = instMarketEye.GetHeaderValue(2)

# GetData
sumPer = 0
for i in range(numStock):
    sumPer += instMarketEye.GetDataValue(0, i)

print("Average PER: ", sumPer / numStock)