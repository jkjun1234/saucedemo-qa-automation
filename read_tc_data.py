"""역할: 엑셀 A2의 TC 번호를 출력해 보는 초기 학습용 스크립트입니다.

사용자가 AI 예시를 참고해 작성·실행했던 연습 코드를 보존합니다.
현재 실제 로그인 테스트의 데이터 읽기는 별도 tc_data.py가 담당합니다.
이 파일은 브라우저 테스트나 실행 프로그램의 시작 파일이 아닙니다.
"""

from openpyxl import load_workbook

workbook = load_workbook("SauceDemo_TC_설계서_초안.xlsx", read_only=True)   # 엑셀 파일 읽기 전용으로 open
sheet = workbook["자동화_데이터"]   # 사용할 시트 선택

print(sheet["A2"].value)    # A2 셀 값 가져오기

workbook.close()    # 파일 닫기
