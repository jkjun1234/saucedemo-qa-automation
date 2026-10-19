# TC01 엑셀 연동 — AI 실행 기록

- 실행일: 2026-10-19, 한국 시간 22:29 결과 확인.
- 코드 작성·실행: AI(Codex). 사용자의 직접 실행으로 계산하지 않습니다.
- 실행 환경: Windows, Python 3.12.14, pytest 9.1.1, Playwright 1.63.0, Chromium, 기존 프로젝트 가상환경.
- 원본 TC 엑셀과 사용자가 작성했던 `read_tc_data.py`는 수정하지 않았습니다.

## 범위와 결과

| 구분 | 검증 대상 | 결과 |
|---|---|---|
| 업무 시나리오 TC01 | 엑셀 TC01의 계정으로 로그인 후 예상 URL 일치·최소 한 상품 표시 | Pass |
| 보조 검증 | TC01 행이 두 번째 데이터 행에 있어도 TC-ID로 올바른 세 값 읽기 | Pass |
| 보조 검증 | 없는 TC99 요청 시 다른 행을 대신 쓰지 않고 ValueError로 중단 | Pass |

전체 출력은 [pytest.log](pytest.log), 기계 판독 결과는 [pytest.xml](pytest.xml)에 보관했습니다. `3 passed in 7.01s`는 TC01 한 개와 보조 검증 두 개의 합계입니다. TC04·TC05·TC02·TC03의 실행 결과가 아닙니다.

## 변경 내용

1. `tc_data.py`에서 원본 엑셀을 읽기 전용으로 열고 A열 TC-ID로 행을 찾습니다.
2. B·C·D열을 username·password·expected_url로 반환합니다. 엑셀은 읽은 뒤 닫습니다.
3. 기존 TC01은 입력값·URL 비교값을 이 함수에서 받고, 접속·입력·클릭·상품 표시 검증은 기존 절차를 유지합니다.

## 실행 명령

프로젝트 루트의 PowerShell에서 실행했습니다.

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH = Join-Path (Get-Location).Path '.browsers'
.\.venv\Scripts\python.exe -m pytest -v --tb=short --tracing=retain-on-failure --basetemp=.pytest_cache/tmp-tc01-full-20261019 --junitxml=evidence/20261019_tc01_excel_ai/pytest.xml
```

위 명령의 `--basetemp`는 이 AI 실행 환경의 임시 폴더 문제를 피하기 위한 지정입니다. 재실행 증거는 별도 날짜·폴더로 저장해 기존 결과를 보존합니다. pytest는 지정한 basetemp 내부를 정리하므로 TC나 보관 증거가 있는 폴더를 지정하지 않습니다.

## 구현 중 확인한 실패와 한계

- 최초 보조 테스트는 pytest 기본 임시 폴더 생성 단계에서 OSError로 중단됐습니다. 프로젝트 내 캐시 폴더로 임시 위치를 지정하자 실행됐습니다. 제품 결함이 아닌 AI 실행 환경 문제입니다.
- 읽기 함수 작성 전 보조 테스트 두 개가 `ModuleNotFoundError: No module named 'tc_data'`로 실패했고, 구현 후 `2 passed in 0.53s`를 확인했습니다. 이는 구현 전 확인이며 계획한 모호한 버튼 선택 오류 분석 사례와 별개입니다.
- 변경 전 TC01 실행은 제한된 실행 환경에서 출력 없이 대기하여 중단했습니다. 변경 전 통과로 기록하지 않습니다. 변경 후 전체 실행은 브라우저 실행이 가능한 환경에서 실제 사이트를 대상으로 완료했습니다.
- 사용자 재실행·코드 설명 확인, 결과 엑셀 자동 저장, 의도적 버튼 선택 오류의 Trace 분석·개선 사례는 아직 완료되지 않았습니다.

## 참고한 공식 문서

- [openpyxl 읽기 전용 로딩·행 순회](https://openpyxl.readthedocs.io/en/stable/tutorial.html)
- [Playwright Python 테스트 작성](https://playwright.dev/python/docs/writing-tests)
