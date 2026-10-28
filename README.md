# SauceDemo 자동화 테스트 학습 포트폴리오

Python · Playwright · pytest · Chromium으로 웹 테스트를 작성·실행하고 결과를 분석하는 학습 프로젝트입니다. 2026-10-19부터 시간 제약에 따라 AI가 기능 단위로 작성·검증하고 사용자가 검토하는 방식으로 진행합니다. 실제 수행자를 구분해 기록합니다.

## Python 파일 역할

각 Python 파일의 최상단에도 아래 역할을 설명했습니다.

| 파일 | 역할 |
|---|---|
| `excel_io.py` | 실행용 엑셀의 TC·단계를 읽고 실행 준비 오류 검사. 브라우저 조작·결과 쓰기는 아직 없음 |
| `tools/prepare_tc_workbook.py` | 기존 초안을 보존한 최초 실행용 양식 준비. 기존 실행용 파일은 덮어쓰지 않음 |
| `tests/e2e/test_tc01_login.py` | 이전 Python 고정 절차 방식의 실제 SauceDemo 로그인 검사. 새 단계 실행기로 수행한 결과가 아님 |
| `tc_data.py` | 이전 방식에서 TC 번호에 맞는 로그인 입력값·비교값 읽기 |
| `tests/unit/` | 임시 엑셀을 이용한 개발용 검사. 사이트 접속 없음 |
| `read_tc_data.py` | 사용자가 작성했던 A2 출력 학습 코드 보존 |

## 현재 구현 상태

- 2026-10-28 작업 1: [실행용 엑셀](SauceDemo_TC_실행용_v2.xlsx)을 준비했습니다. 사람용 TC는 번호순이며 `실제 결과` 열을 제거했고 기존 `결과`·`비고`를 유지했습니다. 결과는 실행 전이므로 빈칸입니다. `자동화_실행단계`에는 TC01 7단계만 준비되어 있습니다. TC02~05는 사람용 초안이며 아직 자동화 단계가 없습니다.
- `excel_io.read_cases()`는 열 이름으로 데이터를 찾아 단계가 준비된 TC만 반환합니다. 잘못된 동작·순서·대상·기대값·검증 누락은 시트/행 사유와 함께 중단합니다. 과거 Pass/Fail은 실행 입력으로 반환하지 않습니다.
- 새 단계 실행기·프로그램 창·결과 자동 갱신은 아직 미구현입니다. 완성 시 선택한 TC 파일의 기존 `결과`에 Pass/Fail/N/A, Fail·N/A의 사유·증거를 `비고`에 기록하고 결과 사본은 만들지 않습니다.
- 실행용 v2는 양식 전환 때 한 번 준비한 입력 파일입니다. 실행마다 만드는 결과 사본이 아닙니다. 기존 초안은 이전 사용자 기록 보존용입니다. 준비 도구는 이미 존재하는 실행용 파일을 덮어쓰지 않습니다.

- 목표: 정상 로그인·잠긴 계정·장바구니 추가/삭제·우편번호 누락·정상 주문의 5개 TC를 작성하고 실행·분석·개선 과정을 남깁니다. 작성·검토 순서는 TC01 → TC02 → TC03 → TC04 → TC05입니다. 로그인 관련 TC01·TC02를 먼저 마칩니다.
- 테스트 설계표(TC): [SauceDemo_TC_설계서_초안.xlsx](SauceDemo_TC_설계서_초안.xlsx). 기존 게임 QA TC의 10개 열 구성을 참고했습니다. TC01에는 사용자 수동·자동화 결과가 기록됐고, TC04는 사용자 수동 오류 화면 확인까지 보고됐습니다.
- 엑셀 읽기 연습: [read_tc_data.py](read_tc_data.py). 현재 `자동화_데이터!A2`의 TC 번호를 출력하는 단계이며, 사용자가 실행해 `TC01` 출력이 나왔다고 보고했습니다.
- 사용자 제시 로그인 성공 기준: **"사이트 접속 및 상품 목록들이 뜨는지 확인".** 사용자는 정상 계정으로 직접 로그인해 상품 목록을 확인했다고 보고했습니다. AI의 열린 화면 확인은 별도 관찰입니다.
- [이전 TC01 코드](tests/e2e/test_tc01_login.py)는 사용자가 AI의 단계별 설명을 참고해 직접 입력·수정했습니다. 2026-09-30 첫 실행 `passed`는 사용자 보고이며, 두 번째 실행은 [저장된 결과](tc01_second_run_20260930.txt)에서 `1 passed in 2.13s`를 확인했습니다. URL과 상품 표시 두 조건의 의미는 선택형으로 확인했고, 입력값·선택자 선택 이유를 본인 말로 설명하는 단계는 남아 있습니다. 다음은 새 엑셀 단계 실행기로 TC01 전체 흐름을 연결한 뒤 TC02~05로 확장합니다.
- TC01 엑셀 연동: [tc_data.py](tc_data.py)가 `자동화_데이터` 시트에서 TC-ID로 행을 찾아 아이디·비밀번호·예상 URL을 읽습니다. 로그인 절차와 상품 표시 검증은 기존 Python 코드에 유지했습니다. 예상 URL은 로그인 후 실제 주소와 비교하며, 로그인 없이 직접 접속할 주소로 사용하지 않습니다.
- 2026-10-19 AI 검증: TC01 실제 사이트 테스트 1개와 엑셀 읽기 보조 테스트 2개가 통과했습니다([실행 기록](evidence/20261019_tc01_excel_ai/README.md)). 10/28에는 새 개발용 검사 19개·기존 개발용 2개와 이전 방식 TC01 1개가 전체 회귀 확인에서 통과했습니다([10/28 기록](evidence/20261028_excel_preparation_ai/README.md)). 새 엑셀 단계 방식의 사이트 실행·결과 갱신·실패 분석 실습은 아직 수행하지 않았습니다.

## 첫 수동 확인

1. https://www.saucedemo.com/ 에 접속합니다.
2. 공개 데모 계정 `standard_user` / `secret_sauce`로 로그인합니다. 페이지에 공개된 계정 정보가 달라졌으면 현재 페이지를 확인합니다.
3. 로그인 전후 화면에서 달라진 점을 보고, 성공 판단 기준을 본인 말로 한 문장 적습니다.
4. 접속이나 로그인이 안 되면 관찰한 오류를 전달합니다. 확인되지 않은 내용을 Pass로 적지 않습니다.

초기 TC01은 사용자가 AI 예시와 설명을 참고해 작성했습니다. 2026-10-19부터 AI가 작은 기능 하나씩 작성·테스트·설명하고 커밋합니다. 사용자의 직접 실행·수정·설명은 실제 수행한 경우에만 별도로 기록합니다.

## 실행 환경

AI가 Codex에 포함된 Python 3.12.14를 바탕으로 이 폴더 안의 `.venv`를 준비했습니다. 시스템 Python이나 다른 프로젝트의 패키지는 변경하지 않았습니다. 이 가상환경은 현재 PC의 Codex 런타임에 의존하므로 다른 PC로 폴더를 옮길 때는 Python을 준비한 뒤 재생성합니다.

### Windows 노트북에서 이어서 작업하기

노트북에 Git과 Python 3.12를 준비한 뒤 PowerShell에서 실행합니다. 첫 설치에는 인터넷 연결이 필요합니다. `.venv`와 `.browsers`는 PC마다 새로 만들며 GitHub에서 내려받지 않습니다.

```powershell
git clone https://github.com/jkjun1234/saucedemo-qa-automation.git
Set-Location -LiteralPath '.\saucedemo-qa-automation'
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
$env:PLAYWRIGHT_BROWSERS_PATH = Join-Path (Get-Location).Path '.browsers'
.\.venv\Scripts\python.exe -m playwright install chromium
.\.venv\Scripts\python.exe -m pytest .\tests\e2e\test_tc01_login.py -q --headed --tracing=retain-on-failure
```

다음 작업을 시작하기 전 `git pull`로 최신 변경을 받고, 노트북 작업을 마치면 변경 파일을 확인한 뒤 커밋·푸시합니다. 집 PC에서는 다시 `git pull`한 다음 이어갑니다. 같은 TC 엑셀 파일을 두 PC에서 동시에 수정하지 않습니다. 최신 실행 프로그램 설계와 계획은 `docs/superpowers/`에 있으며 아직 구현된 기능이 아닙니다.

설치 버전: Playwright 1.63.0, pytest 9.1.1, pytest-playwright 0.9.0, openpyxl 3.1.5, Chromium 153.0.8010.12. 전체 Python 의존성은 `requirements.txt`에 고정했습니다. AI의 로컬 브라우저 기동·DOM 읽기 확인은 `environment_check.txt`에 저장했고 `pip check`도 통과했습니다. 이후 SauceDemo TC01은 사용자가 직접 실행했습니다.

브라우저 파일은 `.browsers`에 보관합니다. 테스트 실행 전 사용할 PowerShell 환경 설정은 다음과 같습니다.

```powershell
Set-Location -LiteralPath 'C:\Users\jun\Desktop\취업서류\자동화테스트_포트폴리오'
$env:PLAYWRIGHT_BROWSERS_PATH = Join-Path (Get-Location).Path '.browsers'
.\.venv\Scripts\python.exe --version
```

### 다른 Windows PC에서 환경을 재생성할 때

Python 3.12를 먼저 준비하고 이 프로젝트 폴더에서 아래 순서로 실행합니다. 현재 PC에는 이미 설치됐으므로 다시 할 필요가 없습니다.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
$env:PLAYWRIGHT_BROWSERS_PATH = Join-Path (Get-Location).Path '.browsers'
.\.venv\Scripts\python.exe -m playwright install chromium
```

이전 TC01의 개별 실행 명령입니다. 현재 새 단계 실행기로 실행하는 명령은 아닙니다.

```powershell
.\.venv\Scripts\python.exe -m pytest .\tests\e2e\test_tc01_login.py -q --headed --tracing=retain-on-failure
```

개발용 검사는 아래처럼 별도로 실행합니다. 통과 개수를 실제 업무 TC 완료 건수로 계산하지 않습니다. 임시 엑셀만 쓰며 사용자 TC 파일은 수정하지 않습니다.

```powershell
.\.venv\Scripts\python.exe -m pytest .\tests\unit -q
```

기존 TC의 실제 사이트 확인은 `tests/e2e`로 분리했습니다. 향후 프로그램의 시작 버튼은 개발용 검사를 호출하지 않습니다.

## 기여 구분

- 사용자: 성공 기준 제시, TC01 코드 직접 입력·수정·응용, 수동 확인과 자동화 실행, TC 결과 정정, 엑셀 데이터 작성, 엑셀 읽기 연습 코드 작성·실행.
- AI: 초기 코드 예시·설명·검토, TC 초안, 환경 설치·점검, 문서 정리, Git·공개 저장소·푸시 알림 설정을 지원했습니다. 2026-10-19의 TC01 엑셀 연동 코드·보조 테스트 작성과 해당 실행은 AI가 수행했습니다. 이전 사용자 작성·실행 기록과 구분합니다.
- AI의 환경 점검과 사용자의 테스트 실행을 구분합니다. 아직 수행하지 않은 테스트와 계획한 기능은 완료 실적으로 계산하지 않습니다.

## Git 버전 관리

2026-10-19부터 이 프로젝트 폴더를 독립 Git 저장소로 관리합니다. 테스트 코드·TC 엑셀·설치 목록·안내 문서·보관한 실행 증거를 추적합니다. 가상환경·브라우저 설치 파일·캐시·자동 생성 Trace·개인 일정과 지원 맥락이 담긴 운영 메모는 로컬에 보관하며 Git에서 제외합니다.

아래 명령은 이 프로젝트 폴더에서 사용합니다.

```powershell
git status --short
git diff -- read_tc_data.py
git log --oneline -5
```

코드를 수정하고 실행 결과를 확인한 뒤, 해당 파일만 골라 커밋합니다. 예를 들어 엑셀에서 읽는 항목을 실제로 추가했을 때 다음처럼 기록할 수 있습니다.

```powershell
git add -- read_tc_data.py
git commit -m "feat: read login data from Excel"
```

커밋은 그 시점의 저장본을 남기는 작업입니다. 커밋만으로 테스트 통과나 포트폴리오 완성을 의미하지 않습니다. TC 엑셀은 이진 파일이므로 표의 상세 변경은 엑셀에서 확인합니다.

공개 저장소: https://github.com/jkjun1234/saucedemo-qa-automation . 검증한 커밋을 `git push`하면 연결된 Discord 채널에 알림이 전송됩니다. 로컬 커밋만으로는 알림을 보내지 않습니다. 웹훅 주소는 저장소에 포함하지 않습니다.

## 참고

- [Playwright Python 시작하기](https://playwright.dev/python/docs/intro)
- [테스트 작성](https://playwright.dev/python/docs/writing-tests)
- [요소 선택](https://playwright.dev/python/docs/locators)
- [Trace Viewer](https://playwright.dev/python/docs/trace-viewer-intro)

기존 Selenium 업무 경험과 이번 학습 프로젝트는 별도 사례입니다. VMS·모바일 검증, 성능 개선 수치, 사용자 직접 수행 범위를 추정하지 않습니다.
