"""역할: 실제 Tk 창의 시작·중지·오류·종료 상태를 개발용 가짜 실행기로 검사합니다.

브라우저 판정은 대체하지 않고 화면 제어만 검사합니다. 실제 사이트 실행은
별도 e2e 검사이며, 이 검사를 사용자 직접 조작 실적으로 계산하지 않습니다.
"""

import threading
import time

import pytest


@pytest.fixture(scope="module")
def tk_root():
    """Tk 인터프리터는 하나만 두고 검사 창은 Toplevel로 분리합니다."""
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()
    yield root
    root.destroy()


@pytest.fixture
def window(monkeypatch, tk_root):
    import tkinter as tk
    import app

    root = tk.Toplevel(tk_root)
    root.withdraw()
    ui = app.AutomationApp(root)
    notices = []
    monkeypatch.setattr(app.messagebox, "showerror", lambda title, message: notices.append(message))
    yield ui, notices
    ui.stop_event.set()
    if ui.worker:
        ui.worker.join(timeout=2)
    if not ui.closed:
        ui.close()


def pump(ui, condition):
    deadline = time.monotonic() + 3
    while not condition() and time.monotonic() < deadline:
        ui.root.update()
        time.sleep(0.01)
    assert condition()


def prepare(monkeypatch, tmp_path):
    import app

    path = tmp_path / "tc.xlsx"
    path.touch()
    cases = [{"tc_id": "TC01", "scenario": "로그인", "steps": []}]
    monkeypatch.setattr(app, "read_cases", lambda path: cases)
    monkeypatch.setattr(app, "write_results", lambda source, results: source)
    return path


def test_invalid_file_does_not_start_worker(window):
    ui, notices = window
    ui.file_path.set("없는파일.xlsx")
    ui.start()
    assert not ui.running
    assert ui.worker is None
    assert notices


def test_start_stop_rejects_duplicate_and_resets_for_rerun(window, monkeypatch, tmp_path):
    import app

    ui, _ = window
    path = prepare(monkeypatch, tmp_path)
    calls = []
    def fake_run(cases, base_url, emit, stop_event, *, slow_mo=0):
        calls.append((threading.get_ident(), slow_mo))
        emit({"type": "tc_started", "tc_id": "TC01"})
        emit({"type": "step_started", "tc_id": "TC01", "step": 1, "action": "접속"})
        assert stop_event.wait(2)
        result = {"tc_id": "TC01", "status": "N/A", "reason": "사용자 중지", "evidence": {}}
        emit({"type": "tc_finished", **result})
        return [result]
    monkeypatch.setattr(app, "run_cases", fake_run)
    ui.file_path.set(str(path))
    ui.start()
    ui.start()
    pump(ui, lambda: "1" in ui.current.get())
    assert ui.start_button.instate(["disabled"])
    ui.stop()
    pump(ui, lambda: not ui.running)
    assert len(calls) == 1
    assert calls[0][0] != threading.get_ident()
    assert calls[0][1] == 5000
    assert ui.results.item("TC01", "values")[1] == "N/A"
    assert "N/A 1" in ui.counts.get()
    assert ui.start_button.instate(["!disabled"])
    ui.slow_view.set(False)
    ui.start()
    assert not ui.stop_event.is_set()
    ui.stop()
    pump(ui, lambda: not ui.running)
    assert len(calls) == 2
    assert calls[1][1] == 0


def test_save_error_is_separate_from_pass(window, monkeypatch, tmp_path):
    import app

    ui, notices = window
    path = prepare(monkeypatch, tmp_path)
    result = {"tc_id": "TC01", "status": "Pass", "reason": "", "evidence": {}}
    def fake_run(cases, url, emit, stop_event, **options):
        emit({"type": "tc_finished", **result})
        return [result]
    def locked(*args):
        raise PermissionError("파일 사용 중")
    monkeypatch.setattr(app, "run_cases", fake_run)
    monkeypatch.setattr(app, "write_results", locked)
    ui.file_path.set(str(path))
    ui.start()
    pump(ui, lambda: not ui.running)
    assert ui.results.item("TC01", "values")[1] == "Pass"
    assert "저장 실패" in ui.save_state.get()
    assert "파일 사용 중" in notices[0]


def test_close_requests_stop_and_waits_for_worker(window, monkeypatch, tmp_path):
    import app

    ui, _ = window
    path = prepare(monkeypatch, tmp_path)
    released = threading.Event()
    def fake_run(cases, url, emit, stop_event, **options):
        stop_event.wait(2)
        released.wait(2)
        return [{"tc_id": "TC01", "status": "N/A", "reason": "사용자 중지", "evidence": {}}]
    monkeypatch.setattr(app, "run_cases", fake_run)
    ui.file_path.set(str(path))
    ui.start()
    ui.close()
    assert ui.stop_event.is_set()
    assert not ui.closed
    released.set()
    pump(ui, lambda: ui.closed)
