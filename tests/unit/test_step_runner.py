"""역할: 로컬 연습 화면으로 단계 실행·실패 분리·중지를 검사합니다.

가상 TC는 개발용이며 실제 SauceDemo TC 수행 실적으로 계산하지 않습니다.
"""
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from threading import Event, Thread

import pytest


@pytest.fixture
def local_site():
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(b'''<html><body><input id="name"><button id="login"
                onclick="localStorage.setItem('dirty','yes');location.href='/inventory.html?name='+encodeURIComponent(document.querySelector('#name').value)">Login</button>
                <div class="item">one</div><div class="item">two</div>
                <div id="clean"></div><script>if(!localStorage.getItem('dirty'))
                document.querySelector('#clean').textContent='clean';
                else document.querySelector('#clean').remove();</script></body></html>''')

        def log_message(self, *args):
            pass

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_port}/"
    finally:
        server.shutdown()
        server.server_close()
        thread.join()


def step(order, action, target="", value="", expected=""):
    return dict(order=order, action=action, target=target, value=value,
                expected=expected, description="가상 화면 검사", source_row=order + 1)


def case(tc_id, steps):
    return dict(tc_id=tc_id, scenario="개발용 가상 TC", steps=steps)


def test_performs_steps_and_emits_progress(local_site, tmp_path, run_in_worker):
    run_cases = run_in_worker
    events = []
    cases = [case("TC01", [step(1, "접속"), step(2, "입력", "#name", "hello"),
        step(3, "클릭", "#login"), step(4, "URL확인", expected="/inventory.html?name=hello"),
        step(5, "표시확인", ".item:first-of-type"),
        step(6, "최소개수확인", ".item", expected=1)])]
    results = run_cases(cases, local_site, events.append, Event(),
                        output_dir=tmp_path, headless=True, timeout_ms=600, slow_mo=1)
    assert results[0]["status"] == "Pass"
    assert results[0]["failed_step"] is None
    assert len([e for e in events if e["type"] == "step_finished"]) == 6
    assert events[-1]["type"] == "run_finished"


def test_failure_saves_evidence_and_next_tc_has_clean_state(local_site, tmp_path, run_in_worker):
    run_cases = run_in_worker
    events = []
    cases = [case("TC01", [step(1, "접속"), step(2, "클릭", "#login"),
        step(3, "URL확인", expected="/wrong"), step(4, "클릭", "#never")]),
        case("TC02", [step(1, "접속"), step(2, "표시확인", "#clean")])]
    results = run_cases(cases, local_site, events.append, Event(),
                        output_dir=tmp_path, headless=True, timeout_ms=300)
    assert [r["status"] for r in results] == ["Fail", "Pass"]
    assert results[0]["failed_step"] == 3
    assert results[0]["reason"]
    assert not any(e.get("tc_id") == "TC01" and e.get("step") == 4 for e in events)
    from pathlib import Path
    assert Path(results[0]["evidence"]["screenshot"]).is_file()
    assert Path(results[0]["evidence"]["trace"]).is_file()


def test_stop_between_steps_leaves_current_and_remaining_na(local_site, tmp_path, run_in_worker):
    run_cases = run_in_worker
    stop = Event()
    events = []
    def receive(event):
        events.append(event)
        if event["type"] == "step_finished":
            stop.set()
    cases = [case("TC01", [step(1, "접속"), step(2, "표시확인", ".item")]),
             case("TC02", [step(1, "접속"), step(2, "표시확인", ".item")])]
    results = run_cases(cases, local_site, receive, stop,
                        output_dir=tmp_path, headless=True)
    assert [r["status"] for r in results] == ["N/A", "N/A"]
    assert all("중지" in r["reason"] for r in results)
    assert len([e for e in events if e["type"] == "step_started"]) == 1


def test_rejects_invalid_base_url_before_launch(tmp_path, run_in_worker):
    run_cases = run_in_worker
    with pytest.raises(ValueError, match="http"):
        run_cases([], "saucedemo.com", lambda event: None, Event(), output_dir=tmp_path)


def test_excel_input_and_expected_changes_affect_real_browser(local_site, tmp_path, run_in_worker):
    """엑셀만 바꿔 입력과 판정이 바뀌는지 실제 로컬 화면에서 확인합니다."""
    from openpyxl import Workbook
    from excel_io import read_cases
    run_cases = run_in_worker
    workbook = Workbook()
    tc = workbook.active
    tc.title = "TC_웹주문"
    tc.append(["TC-ID", "테스트 시나리오", "결과", "비고"])
    tc.append(["TC01", "개발용 입력 확인", None, None])
    steps = workbook.create_sheet("자동화_실행단계")
    steps.append(["TC-ID", "순서", "동작", "대상", "입력값", "기대값", "설명"])
    steps.append(["TC01", 1, "접속", None, None, None, None])
    steps.append(["TC01", 2, "입력", "#name", "alice", None, None])
    steps.append(["TC01", 3, "클릭", "#login", None, None, None])
    steps.append(["TC01", 4, "URL확인", None, None, "/inventory.html?name=alice", None])
    path = tmp_path / "editable.xlsx"

    def execute():
        workbook.save(path)
        return run_cases(read_cases(path), local_site, lambda event: None, Event(),
                         output_dir=tmp_path, headless=True, timeout_ms=300)[0]

    try:
        assert execute()["status"] == "Pass"
        steps["E3"] = "bob"
        changed_input = execute()
        assert changed_input["status"] == "Fail"
        assert changed_input["failed_step"] == 4
        steps["F5"] = "/inventory.html?name=bob"
        assert execute()["status"] == "Pass"
    finally:
        workbook.close()


def test_stop_during_element_wait_is_na_not_fail(local_site, tmp_path, run_in_worker):
    from threading import Timer
    stop = Event()
    timer = None
    def receive(event):
        nonlocal timer
        if event["type"] == "step_started" and event["step"] == 2:
            timer = Timer(0.05, stop.set)
            timer.start()
    cases = [case("TC01", [step(1, "접속"), step(2, "클릭", "#missing"),
                           step(3, "표시확인", ".item")])]
    try:
        results = run_in_worker(cases, local_site, receive, stop,
                                output_dir=tmp_path, headless=True, timeout_ms=300)
        assert results[0]["status"] == "N/A"
        assert "중지" in results[0]["reason"]
    finally:
        if timer:
            timer.cancel()
            timer.join()


def test_closed_browser_page_is_environment_error(local_site, tmp_path, run_in_worker, monkeypatch):
    import step_runner
    real_perform = step_runner._perform_step
    def close_then_perform(page, current, *args):
        if current["order"] == 2:
            page.close()
        return real_perform(page, current, *args)
    monkeypatch.setattr(step_runner, "_perform_step", close_then_perform)
    events = []
    cases = [case("TC01", [step(1, "접속"), step(2, "표시확인", ".item")]),
             case("TC02", [step(1, "접속"), step(2, "표시확인", ".item")])]
    results = run_in_worker(cases, local_site, events.append, Event(), output_dir=tmp_path, headless=True)
    assert [r["status"] for r in results] == ["N/A", "N/A"]
    assert any(e["type"] == "run_error" and e["category"] == "환경" for e in events)


@pytest.mark.parametrize("action,target,expected,status", [
    ("문구확인", "#clean", "clean", "Pass"),
    ("문구확인", "#clean", "잘못된 문구", "Fail"),
    ("개수확인", ".missing", 0, "Pass"),
    ("개수확인", ".missing", 1, "Fail"),
])
def test_text_and_exact_count_use_real_page(local_site, tmp_path, run_in_worker, action, target, expected, status):
    cases = [case("TC02", [step(1, "접속"), step(2, action, target, expected=expected)])]
    results = run_in_worker(cases, local_site, lambda event: None, Event(),
                            output_dir=tmp_path, headless=True, timeout_ms=250)
    assert results[0]["status"] == status
    if status == "Fail":
        assert results[0]["failed_step"] == 2
        assert "지원하지 않는" not in results[0]["reason"]
