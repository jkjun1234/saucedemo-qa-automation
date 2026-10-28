"""역할: 새 실행기가 실행용 엑셀 TC01을 실제 SauceDemo에서 수행하는지 확인합니다.

사용자 TC 파일은 읽기 전용으로 사용합니다. 판정·로그는 실행기 반환값과
test-results에 남으며 TC 엑셀에 자동으로 기록하는 기능은 다음 작업입니다.
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
