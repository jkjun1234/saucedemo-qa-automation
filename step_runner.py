"""역할: 엑셀에서 읽은 동작을 Playwright 브라우저에서 순서대로 실행합니다.

excel_io.read_cases()의 반환값을 받으며 TC마다 새 컨텍스트를 만듭니다.
진행 이벤트·판정·실패 증거·실행 로그를 제공하지만 창 표시나 TC 엑셀
결과 수정은 하지 않습니다. 프로그램 창은 향후 run_cases()를 호출합니다.
"""
import json
import re
from decimal import Decimal
from datetime import datetime, timedelta, timezone
from pathlib import Path
from uuid import uuid4
from urllib.parse import urljoin, urlsplit

from playwright.sync_api import Error, TimeoutError, expect, sync_playwright
from excel_io import SAVE_ACTIONS, SAVED_CHECK_ACTIONS


KST = timezone(timedelta(hours=9))


def _now():
    return datetime.now(KST).isoformat(timespec="milliseconds")


def _usd_amount(text):
    """데모의 $숫자.소수2자리 한 금액을 정확한 Decimal 값으로 읽습니다."""
    matches = re.findall(r"\$(\d+\.\d{2})(?![\d.,])", text)
    if len(matches) != 1:
        raise ValueError(f"달러 소수2자리 금액 하나를 읽을 수 없습니다: {text}")
    return Decimal(matches[0])


class _SavedComparisonError(AssertionError):
    """저장값 비교 실패의 기대값·관찰값을 로그와 엑셀에 함께 전달합니다."""
    def __init__(self, key, expected, actual):
        super().__init__(f"{key} 기대값 {expected}, 실제값 {actual}")
        self.detail = {"value_key": key, "expected_value": str(expected), "observed_value": str(actual)}


def _perform_step(page, step, base_url, timeout_ms, saved_values):
    """지원하는 한 동작만 수행하며 임의 코드를 실행하지 않습니다."""
    action = step["action"]
    if action == "접속":
        page.goto(urljoin(base_url, step["value"]), timeout=timeout_ms)
    elif action == "입력":
        page.locator(step["target"]).fill(str(step["value"]))
    elif action == "클릭":
        page.locator(step["target"]).click()
    elif action == "URL확인":
        expect(page).to_have_url(urljoin(base_url, step["expected"]), timeout=timeout_ms)
    elif action == "표시확인":
        expect(page.locator(step["target"])).to_be_visible(timeout=timeout_ms)
    elif action == "문구확인":
        expect(page.locator(step["target"])).to_have_text(step["expected"], timeout=timeout_ms)
    elif action == "개수확인":
        expect(page.locator(step["target"])).to_have_count(step["expected"], timeout=timeout_ms)
    elif action in SAVE_ACTIONS:
        raw = page.locator(step["target"]).inner_text().strip()
        kind = SAVE_ACTIONS[action]
        value = _usd_amount(raw) if kind == "금액" else raw
        saved_values[step["value"]] = (kind, value)
        return {"value_key": step["value"], "stored_value": str(value)}
    elif action in SAVED_CHECK_ACTIONS:
        key = step["value"]
        if key not in saved_values:
            raise ValueError(f"같은 TC에서 먼저 저장하지 않은 값: {key}")
        kind, expected = saved_values[key]
        if kind != SAVED_CHECK_ACTIONS[action]:
            raise ValueError(f"저장값 종류가 다릅니다: {key}")
        locator = page.locator(step["target"])
        if kind == "문구":
            try:
                expect(locator).to_have_text(expected, timeout=timeout_ms)
            except AssertionError as error:
                actual = " | ".join(locator.all_text_contents()) or "(요소 없음)"
                raise _SavedComparisonError(key, expected, actual) from error
            actual = locator.inner_text().strip()
        else:
            actual = _usd_amount(locator.inner_text())
            if actual != expected:
                raise _SavedComparisonError(key, expected, actual)
        return {"value_key": key, "expected_value": str(expected), "observed_value": str(actual)}
    elif action == "최소개수확인":
        try:
            page.wait_for_function(
                "args => document.querySelectorAll(args.selector).length >= args.minimum",
                arg={"selector": step["target"], "minimum": step["expected"]},
                timeout=timeout_ms,
            )
        except TimeoutError as error:
            count = page.locator(step["target"]).count()
            raise AssertionError(f"최소 {step['expected']}개 기대, 실제 {count}개") from error
    else:
        raise ValueError(f"지원하지 않는 동작: {action}")


def run_cases(cases, base_url, emit, stop_event, *, output_dir=None,
              headless=False, timeout_ms=10_000, slow_mo=0):
    """준비된 TC를 실행하고 결과 목록을 반환합니다. 자동 TC 재시도는 없습니다.

    emit은 진행 이벤트를 받는 함수, stop_event는 threading.Event입니다.
    중지는 현재 동작이 끝나거나 제한시간에 도달한 뒤 단계 경계에서 처리합니다.
    output_dir는 실행별 증거 폴더이며 기존 실행 결과를 덮어쓰지 않습니다.
    GUI/pytest에서는 기존 이벤트 루프와 겹치지 않는 작업 스레드에서 호출합니다.
    slow_mo는 브라우저 동작 사이의 관찰용 간격(ms)이며 판정 조건은 바꾸지 않습니다.
    """
    parsed = urlsplit(base_url.strip())
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("대상 주소에 http:// 또는 https://와 호스트를 포함하세요.")
    base_url = base_url.strip().rstrip("/") + "/"
    run_id = datetime.now(KST).strftime("%Y%m%d_%H%M%S") + "_" + uuid4().hex[:8]
    folder = Path(output_dir or Path(__file__).resolve().parent / "test-results") / run_id
    folder.mkdir(parents=True, exist_ok=False)
    log_path = folder / "events.jsonl"
    results = [dict(tc_id=c["tc_id"], status="N/A", expected="모든 단계의 검증 통과",
                    actual="미실행", failed_step=None, reason="실행 전",
                    started_at=None, finished_at=None, evidence={"log": str(log_path)})
               for c in cases]

    with log_path.open("w", encoding="utf-8") as log:
        def send(event_type, **fields):
            event = {"type": event_type, "time": _now(), "run_id": run_id, **fields}
            log.write(json.dumps(event, ensure_ascii=False) + "\n")
            log.flush()
            emit(event)

        send("run_started", total=len(cases), base_url=base_url)
        try:
            if cases and not stop_event.is_set():
                with sync_playwright() as playwright:
                    browser = playwright.chromium.launch(headless=headless, slow_mo=slow_mo)
                    try:
                        for case, result in zip(cases, results):
                            if stop_event.is_set():
                                result["reason"] = "사용자 중지로 미실행"
                                result["finished_at"] = _now()
                                send("tc_finished", **result)
                                continue
                            result["started_at"] = _now()
                            send("tc_started", tc_id=case["tc_id"])
                            saved_values = {}
                            context = browser.new_context()
                            try:
                                context.tracing.start(screenshots=True, snapshots=True, sources=True)
                                page = context.new_page()
                                page.set_default_timeout(timeout_ms)
                                current = None
                                try:
                                    for current in case["steps"]:
                                        if stop_event.is_set():
                                            result["reason"] = "사용자 중지로 TC 미완료"
                                            break
                                        fields = dict(tc_id=case["tc_id"], step=current["order"],
                                                      action=current["action"], target=current["target"])
                                        send("step_started", **fields)
                                        detail = _perform_step(page, current, base_url, timeout_ms, saved_values)
                                        send("step_finished", **fields, **(detail or {}))
                                    else:
                                        result.update(status="Pass", actual="모든 단계와 검증 통과", reason="")
                                except (AssertionError, Error, ValueError) as error:
                                    if stop_event.is_set():
                                        result.update(actual="대기 중 중지되어 미완료", reason="사용자 중지로 TC 미완료")
                                        send("step_stopped", tc_id=case["tc_id"], step=current["order"])
                                    elif page.is_closed() or not browser.is_connected():
                                        # 닫힌 브라우저의 오류는 업무 Fail이 아닌 전체 환경 오류로 전달합니다.
                                        raise
                                    else:
                                        result.update(status="Fail", actual=str(error), reason=str(error),
                                                      failed_step=current["order"] if current else None)
                                        detail = error.detail if isinstance(error, _SavedComparisonError) else {}
                                        if detail:
                                            result.update(expected=detail["expected_value"], actual=detail["observed_value"])
                                        send("step_failed", tc_id=case["tc_id"], step=result["failed_step"],
                                             category="검증" if isinstance(error, AssertionError) else "단계 실행",
                                             message=str(error), **detail)
                                        screenshot = folder / f"{case['tc_id']}_failure.png"
                                        try:
                                            page.screenshot(path=str(screenshot), full_page=True)
                                            result["evidence"]["screenshot"] = str(screenshot)
                                        except Error as evidence_error:
                                            send("evidence_error", tc_id=case["tc_id"], message=str(evidence_error))
                                if result["status"] == "Fail":
                                    trace = folder / f"{case['tc_id']}_trace.zip"
                                    try:
                                        context.tracing.stop(path=str(trace))
                                        result["evidence"]["trace"] = str(trace)
                                    except Error as evidence_error:
                                        send("evidence_error", tc_id=case["tc_id"], message=str(evidence_error))
                                else:
                                    context.tracing.stop()
                            finally:
                                context.close()
                            result["finished_at"] = _now()
                            send("tc_finished", **result)
                    finally:
                        browser.close()
            elif stop_event.is_set():
                for result in results:
                    result.update(reason="사용자 중지로 미실행", finished_at=_now())
                    send("tc_finished", **result)
        except Error as error:
            send("run_error", category="환경", message=str(error))
            for result in results:
                if result["finished_at"] is None:
                    result.update(reason=f"환경 오류로 실행 중단: {error}", finished_at=_now())
                    send("tc_finished", **result)
        finally:
            send("run_finished", results=results)
    return results
