"""역할: 엑셀 실행 단계 읽기 기능을 임시 파일로 검사합니다.

실제 SauceDemo나 업무 TC를 실행하지 않는 개발용 검사입니다.
각 검사는 임시 엑셀을 사용하며 사용자의 TC 파일은 수정하지 않습니다.
"""

from hashlib import sha256

import pytest
from openpyxl import Workbook


TC_HEADERS = ["TC-ID", "테스트 시나리오", "결과", "비고"]
STEP_HEADERS = ["TC-ID", "순서", "동작", "대상", "입력값", "기대값", "설명"]


@pytest.fixture
def make_workbook(tmp_path):
    """원본과 별개인 작은 TC·실행 단계 표를 만들어 검사에 제공합니다."""
    def make(tc_rows=None, steps=None, headers=None):
        workbook = Workbook()
        tc = workbook.active
        tc.title = "TC_웹주문"
        tc.append(["임시 검사 데이터"])
        tc.append(TC_HEADERS)
        for row in tc_rows or [["TC01", "가상 로그인", "Pass", "과거 기록"]]:
            tc.append(row)
        step_sheet = workbook.create_sheet("자동화_실행단계")
        step_sheet.append(STEP_HEADERS if headers is None else headers)
        default_steps = [
            ["TC01", 1, "접속", None, None, None, "첫 화면"],
            ["TC01", 2, "URL확인", None, None, "/inventory.html", "이동 확인"],
        ]
        for row in default_steps if steps is None else steps:
            step_sheet.append(row)
        path = tmp_path / "cases.xlsx"
        workbook.save(path)
        workbook.close()
        return path
    return make


def test_reads_cases_by_id_and_orders_steps_without_old_results(make_workbook):
    """엑셀 행 순서가 아닌 TC 번호·단계 번호로 읽고 과거 Pass는 넘기지 않습니다."""
    from excel_io import read_cases

    path = make_workbook(
        tc_rows=[["TC10", "가상 열 번째", "Pass", "옛 결과"],
                 ["TC01", "가상 첫 번째", "Fail", "옛 사유"],
                 ["TC02", "단계 미준비", None, None]],
        steps=[["TC10", 1, "URL확인", None, None, "/ten", None],
               ["TC01", 2, "URL확인", None, None, "/one", None],
               ["TC01", 1, "접속", None, None, None, None]],
    )
    cases = read_cases(path)
    assert [case["tc_id"] for case in cases] == ["TC01", "TC10"]
    assert cases[0]["scenario"] == "가상 첫 번째"
    assert [step["order"] for step in cases[0]["steps"]] == [1, 2]
    assert cases[0]["steps"][1]["expected"] == "/one"
    assert set(cases[0]) == {"tc_id", "scenario", "steps"}


def test_preserves_blank_input_and_password_spaces(make_workbook):
    """빈 입력은 지우기 동작에 쓰이며 비밀번호의 앞뒤 공백을 임의 제거하면 안 됩니다."""
    from excel_io import read_cases

    path = make_workbook(steps=[
        ["TC01", 1, "입력", "#password", " secret ", None, None],
        ["TC01", 2, "입력", "#postal", None, None, None],
        ["TC01", 3, "최소개수확인", ".item", None, 1, None],
    ])
    steps = read_cases(path)[0]["steps"]
    assert steps[0]["value"] == " secret "
    assert steps[1]["value"] == ""
    assert steps[2]["expected"] == 1


def test_reading_keeps_workbook_bytes_unchanged(make_workbook):
    from excel_io import read_cases

    path = make_workbook()
    before = sha256(path.read_bytes()).hexdigest()
    read_cases(path)
    assert sha256(path.read_bytes()).hexdigest() == before


def test_matches_headers_instead_of_column_letters(make_workbook):
    from excel_io import read_cases

    path = make_workbook(
        headers=["설명", "기대값", "입력값", "대상", "동작", "순서", "TC-ID"],
        steps=[["비교", "/inventory.html", None, None, "URL확인", 1, "TC01"]],
    )
    assert read_cases(path)[0]["steps"][0]["expected"] == "/inventory.html"


@pytest.mark.parametrize(("steps", "error"), [
    ([["TC01", 1, "임의코드실행", None, None, None, None]], "지원하지 않는 동작"),
    ([["TC01", 1, "접속", None, None, None, None]], "검증 단계"),
    ([["TC01", 1, "클릭", None, None, None, None]], "대상"),
    ([["TC01", 1, "URL확인", None, None, None, None]], "기대값"),
    ([["TC01", 1, "접속", None, "https://other.example/", None, None]], "상대 경로"),
    ([["TC01", 1, "최소개수확인", ".item", None, 0, None]], "양의 정수"),
    ([["TC01", 1, "URL확인", None, None, "/", None],
      ["TC01", 1, "접속", None, None, None, None]], "중복"),
    ([["TC01", 1, "접속", None, None, None, None],
      ["TC01", 3, "URL확인", None, None, "/", None]], "연속"),
    ([["TC99", 1, "URL확인", None, None, "/", None]], "TC 표에 없는"),
])
def test_invalid_steps_report_sheet_and_row(make_workbook, steps, error):
    from excel_io import read_cases

    with pytest.raises(ValueError, match=error) as raised:
        read_cases(make_workbook(steps=steps))
    assert "자동화_실행단계" in str(raised.value)
    assert "행" in str(raised.value)


def test_missing_step_header_reports_header_row(make_workbook):
    from excel_io import read_cases

    path = make_workbook(headers=STEP_HEADERS[:-1])
    with pytest.raises(ValueError, match="설명") as raised:
        read_cases(path)
    assert "1행" in str(raised.value)


def test_duplicate_tc_id_is_not_silently_overwritten(make_workbook):
    from excel_io import read_cases

    path = make_workbook(tc_rows=[["TC01", "처음", None, None], ["TC01", "중복", None, None]])
    with pytest.raises(ValueError, match="중복") as raised:
        read_cases(path)
    assert "TC_웹주문" in str(raised.value)


def test_empty_steps_are_not_reported_as_success(make_workbook):
    from excel_io import read_cases

    with pytest.raises(ValueError, match="실행 단계가 없습니다"):
        read_cases(make_workbook(steps=[]))


@pytest.mark.parametrize("action,target,expected", [
    ("문구확인", "[data-test=error]", "잠긴 계정"),
    ("개수확인", ".inventory_item", 0),
])
def test_reads_text_and_exact_count_checks(make_workbook, action, target, expected):
    from excel_io import read_cases

    cases = read_cases(make_workbook(steps=[["TC01", 1, action, target, None, expected, "검증"]]))
    assert cases[0]["steps"][0]["expected"] == expected


@pytest.mark.parametrize("action,target,expected,message", [
    ("문구확인", "#error", None, "기대값"),
    ("문구확인", None, "오류", "대상"),
    ("개수확인", ".item", -1, "0 이상의 정수"),
    ("개수확인", ".item", 1.5, "0 이상의 정수"),
    ("개수확인", ".item", None, "0 이상의 정수"),
])
def test_rejects_missing_text_and_invalid_exact_counts(make_workbook, action, target, expected, message):
    from excel_io import read_cases

    with pytest.raises(ValueError, match=message):
        read_cases(make_workbook(steps=[["TC01", 1, action, target, None, expected, "잘못된 검증"]]))
