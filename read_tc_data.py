from openpyxl import load_workbook

workbook = load_workbook("SauceDemo_TC_설계서_초안.xlsx", read_only=True)   # 엑셀 파일 읽기 전용으로 open
sheet = workbook["자동화_데이터"]   # 사용할 시트 선택

print(sheet["A2"].value)    # A2 셀 값 가져오기

workbook.close()    # 파일 닫기