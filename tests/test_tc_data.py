import pytest
from openpyxl import Workbook


@pytest.fixture
def login_workbook(tmp_path):
    path = tmp_path / "login_data.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "자동화_데이터"
    sheet.append(["TC-ID", "아이디", "비밀번호", "예상 URL"])
    sheet.append(["TC02", "locked_out_user", "other_password", "https://example.com/locked"])
    sheet.append(["TC01", "standard_user", "secret_sauce", "https://www.saucedemo.com/inventory.html"])
    workbook.save(path)
    workbook.close()
    return path


def test_reads_login_values_by_tc_id_instead_of_row_number(login_workbook):
    from tc_data import load_login_data

    assert load_login_data("TC01", login_workbook) == {
        "username": "standard_user",
        "password": "secret_sauce",
        "expected_url": "https://www.saucedemo.com/inventory.html",
    }


def test_missing_tc_id_stops_instead_of_using_another_row(login_workbook):
    from tc_data import load_login_data

    with pytest.raises(ValueError, match="TC99"):
        load_login_data("TC99", login_workbook)
