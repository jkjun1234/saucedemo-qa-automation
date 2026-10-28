# 실행용 TC 양식·단계 읽기 — AI 검증 기록

- 날짜: 2026-10-28, 한국 시간 17:02 전체 출력 확인.
- 수행자: AI(Codex). 사용자 직접 작성·실행으로 계산하지 않습니다.
- 대상 범위: 승인 계획 작업 1. GUI·엑셀 단계의 브라우저 실행·결과 쓰기는 미구현입니다.

## 산출물

- `excel_io.py`: TC-ID와 헤더 이름으로 시나리오·실행 단계를 읽습니다. 실행 단계가 있는 TC만 반환하며 과거 결과·비고는 반환하지 않습니다.
- `SauceDemo_TC_실행용_v2.xlsx`: TC01~05 번호순, 실제 결과 열 제거, 결과·비고 유지 및 초기화, TC01 실행 단계 7개. TC02~05는 단계 미준비입니다.
- `tools/prepare_tc_workbook.py`: 최초 실행용 양식 준비 도구입니다. 기존 초안을 변경하지 않고, 이미 존재하는 실행용 파일도 덮어쓰지 않습니다. 매 실행의 결과 사본 생성 기능이 아닙니다.
- 개발용 검사는 `tests/unit`, 기존 실제 사이트 TC01은 `tests/e2e`로 분리했습니다.

## 검증 결과

| 구분 | 내용 | 결과 |
|---|---|---|
| 새 개발용 검사 | 행/열 순서, TC-ID, 입력 공백 보존, 동작·순서·검증·필수값 오류, 읽기 전용 보존 | 16개 통과 |
| 새 양식 준비 검사 | 초안 해시·TC 내용·데이터 시트·헤더 서식 보존, 결과 열/드롭다운/조건부 서식 범위, 준비 파일 덮어쓰기 차단 | 3개 통과 |
| 기존 개발용 검사 | 이전 로그인 데이터 읽기 | 2개 통과 |
| 기존 업무 TC 회귀 확인 | 이전 Python 고정 절차의 실제 SauceDemo TC01 | 1개 통과 |

전체 출력은 [pytest.log](pytest.log), 기계 판독 결과는 [pytest.xml](pytest.xml)에 보관했습니다. `22 passed in 8.42s`는 개발용 21개와 기존 업무 TC01 1개의 합계입니다. 새 단계 실행기로 TC01~05를 실행한 실적이 아닙니다.

## 방법과 한계

- 신규 읽기 검사는 구현 전 모듈 부재로 16개 실패, 구현 후 통과를 확인했습니다. 준비 도구 검사도 구현 전 3개 실패를 확인했습니다.
- 양식의 헤더 폰트 비교에서는 openpyxl StyleProxy끼리 직접 비교해 동일 서식을 다르게 판정했습니다. 실제 폰트 값과 `copy(font)` 비교가 일치함을 확인하여 검사 비교 방식만 수정했습니다. 제품 결함·실패 분석 포트폴리오 사례로 쓰지 않습니다.
- 기존 초안의 바이트가 보존되는지 검사했습니다. 새 양식에는 과거 Pass를 복사하지 않았습니다. 이번 검증의 Pass는 로그에만 기록했으며 실행용 TC의 결과·비고는 비어 있습니다.
- 렌더 미리보기로 번호순 TC 표, 단계 표, 기존 데이터 시트의 표시를 확인했습니다. 데이터 유효성 목록과 조건부 서식의 적용 범위·기존 3개 규칙을 파일 구조에서 확인했습니다. 실제 Excel/Calc에서 드롭다운 선택·조건부 글자색 반응을 조작해 검증한 것은 아닙니다.
- GUI·단계 실행기·원본 TC 파일 결과 갱신·사용자 직접 실행은 아직 확인하지 않았습니다. 다음 작업은 준비된 TC01 단계를 브라우저에서 실행하는 기능입니다.

## 실행 명령

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH = Join-Path (Get-Location).Path '.browsers'
.\.venv\Scripts\python.exe -m pytest tests -v --tb=short --tracing=retain-on-failure --basetemp=.pytest_cache/tmp-stage1-regression-20261028
```

위 basetemp는 AI 검증용 임시 경로입니다. 데이터·증거 보관 폴더를 basetemp로 지정하지 않습니다. 개발용 검사만 확인할 때는 `python -m pytest tests/unit -q`, 기존 사이트 TC01만 확인할 때는 `python -m pytest tests/e2e/test_tc01_login.py -q`로 구분합니다.

참고: [openpyxl 행/열 변경 시 의존성 처리 범위](https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html), [읽기 전용 로딩과 행 순회](https://openpyxl.readthedocs.io/en/stable/tutorial.html).
