"""역할: 엑셀 읽기 모듈 tc_data.py를 검사하는 보조 테스트입니다.

SauceDemo에 접속하거나 로그인하지 않습니다. 임시 엑셀을 만들어
TC 번호로 올바른 행을 찾는지, 없는 번호를 요청하면 중단하는지 검사합니다.
실제 TC 원본은 수정하지 않으며, 두 개가 통과해도 업무 시나리오 두 개를
실행한 것으로 계산하지 않습니다. 작성·검증: AI, 2026-10-19.
"""

import pytest
from openpyxl import Workbook


@pytest.fixture
def login_workbook(tmp_path):
    """각 검사에서 사용할 임시 엑셀을 만듭니다. tmp_path는 pytest가 제공합니다."""
    path = tmp_path / "login_data.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "자동화_데이터"
    sheet.append(["TC-ID", "아이디", "비밀번호", "예상 URL"])
    # 아래 TC02는 행 순서 검사용 가상 데이터이며 실제 잠긴 계정 TC 실행이 아닙니다.
    sheet.append(["TC02", "locked_out_user", "other_password", "https://example.com/locked"])
    # TC01을 첫 데이터 행이 아닌 곳에 둬서 '무조건 2행 읽기' 오류를 잡습니다.
    sheet.append(["TC01", "standard_user", "secret_sauce", "https://www.saucedemo.com/inventory.html"])
    workbook.save(path)
    workbook.close()
    return path


def test_reads_login_values_by_tc_id_instead_of_row_number(login_workbook):
    """행 위치와 관계없이 TC01에 해당하는 입력값·비교값을 읽어야 합니다."""
    from tc_data import load_login_data

    # 읽은 세 값이 미리 정한 정답과 모두 일치해야 이 검사가 통과합니다.
    assert load_login_data("TC01", login_workbook) == {
        "username": "standard_user",
        "password": "secret_sauce",
        "expected_url": "https://www.saucedemo.com/inventory.html",
    }


def test_missing_tc_id_stops_instead_of_using_another_row(login_workbook):
    """없는 TC99를 요청했을 때 다른 TC 데이터로 실행하는 일을 막습니다."""
    from tc_data import load_login_data

    # 예상한 오류가 발생하면 Pass입니다. 오류 없이 진행하면 이 검사는 Fail입니다.
    with pytest.raises(ValueError, match="TC99"):
        load_login_data("TC99", login_workbook)
