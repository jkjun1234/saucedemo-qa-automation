"""역할: 새 동기 실행기를 프로그램과 같은 별도 작업 스레드에서 검사합니다.

pytest-playwright의 기존 세션과 새 sync_playwright 세션을 같은 스레드에
중첩하지 않습니다. 실행기 동작을 대체하거나 판정을 모의 처리하지 않습니다.
"""
from concurrent.futures import ThreadPoolExecutor

import pytest


@pytest.fixture
def run_in_worker():
    def execute(*args, **kwargs):
        from step_runner import run_cases
        with ThreadPoolExecutor(max_workers=1) as worker:
            return worker.submit(run_cases, *args, **kwargs).result()
    return execute
