# 매수/매도 API는 주식 시장 운영 시간에 코드 실행
import win32com.client

# 객체 생성
instCpTdUtil = win32com.client.Dispatch("CpTrade.CpTdUtil")
instCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311")

# 해당 객체의 TradeInit 메소드 호출 -> 주문을 위한 초기화
instCpTdUtil.TradeInit()

# 주문 종류 코드와 계좌 번호 설정
accountNumber = instCpTdUtil.AccountNumber[0]
instCpTd0311.SetInputValue(0, 2)
instCpTd0311.SetInputValue(1, accountNumber)

# 주문할 종목의 종목 코드 설정
instCpTd0311.SetInputValue(3, "A005930")

# 주문 수량과 주문 단가 입력
instCpTd0311.SetInputValue(4, 10)
instCpTd0311.SetInputValue(5, 80000)

instCpTd0311.BlockRequest()