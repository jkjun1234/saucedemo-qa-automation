from pathlib import Path

from openpyxl import load_workbook


TC_WORKBOOK = Path(__file__).resolve().parent / "SauceDemo_TC_설계서_초안.xlsx"


def load_login_data(tc_id, workbook_path=TC_WORKBOOK):
    """자동화_데이터 시트에서 TC 번호에 해당하는 로그인 값을 읽습니다."""
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        sheet = workbook["자동화_데이터"]
        for row in sheet.iter_rows(min_row=2, max_col=4, values_only=True):
            row_tc_id, username, password, expected_url = row
            if row_tc_id == tc_id:
                return {
                    "username": username,
                    "password": password,
                    "expected_url": expected_url,
                }
        raise ValueError(f"자동화_데이터 시트에서 {tc_id}를 찾지 못했습니다.")
    finally:
        workbook.close()
