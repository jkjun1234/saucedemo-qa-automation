"""역할: 실행용 엑셀의 TC01·TC02를 실제 SauceDemo에서 수행하는지 확인합니다.

사용자 TC 파일은 읽기 전용으로 사용합니다. 판정·로그는 실행기 반환값과
test-results에 남습니다. 이 개발 검사는 프로그램 창의 결과 저장을 호출하지 않습니다.
"""
from pathlib import Path
from threading import Event

from excel_io import read_cases


def test_excel_tc01_on_saucedemo(run_in_worker):
    root = Path(__file__).resolve().parents[2]
    events = []
    cases = [case for case in read_cases(root / "SauceDemo_TC_실행용_v2.xlsx")
             if case["tc_id"] == "TC01"]
    results = run_in_worker(cases,
                        "https://www.saucedemo.com/", events.append, Event(), headless=True)
    assert len(results) == 1
    assert results[0]["tc_id"] == "TC01"
    assert results[0]["status"] == "Pass", results[0]["reason"]
    assert events[-1]["type"] == "run_finished"


def test_excel_tc02_locked_account_on_saucedemo(run_in_worker):
    root = Path(__file__).resolve().parents[2]
    cases = [case for case in read_cases(root / "SauceDemo_TC_실행용_v2.xlsx")
             if case["tc_id"] == "TC02"]
    assert len(cases) == 1, "TC02 실행 단계가 필요합니다."
    assert any(s["action"] == "입력" and s["value"] == "locked_out_user" for s in cases[0]["steps"])
    results = run_in_worker(cases, "https://www.saucedemo.com/", lambda event: None, Event(), headless=True)
    assert results[0]["status"] == "Pass", results[0]["reason"]
