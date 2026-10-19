# SauceDemo 자동화 테스트 학습 포트폴리오

Python · Playwright · pytest · Chromium으로 웹 테스트를 직접 작성하며 배우는 프로젝트입니다.

## 현재 단계

- 목표: 정상 로그인·잠긴 계정·장바구니 추가/삭제·우편번호 누락·정상 주문의 5개 TC를 작성하고 실행·분석·개선 과정을 남깁니다. 우선 작성 순서는 TC01 → TC04 → TC05입니다.
- 테스트 설계표(TC): [SauceDemo_TC_설계서_초안.xlsx](SauceDemo_TC_설계서_초안.xlsx). 기존 게임 QA TC의 10개 열 구성을 참고했습니다. TC01에는 사용자 수동·자동화 결과가 기록됐고, TC04는 사용자 수동 오류 화면 확인까지 보고됐습니다.
- 엑셀 읽기 연습: [read_tc_data.py](read_tc_data.py). 현재 `자동화_데이터!A2`의 TC 번호를 출력하는 단계이며, 사용자가 실행해 `TC01` 출력이 나왔다고 보고했습니다.
- 사용자 제시 로그인 성공 기준: **"사이트 접속 및 상품 목록들이 뜨는지 확인".** 사용자는 정상 계정으로 직접 로그인해 상품 목록을 확인했다고 보고했습니다. AI의 열린 화면 확인은 별도 관찰입니다.
- [TC01 코드](tests/test_tc01_login.py)는 사용자가 AI의 단계별 설명을 참고해 직접 입력·수정했습니다. 2026-09-30 첫 실행 `passed`는 사용자 보고이며, 두 번째 실행은 [저장된 결과](tc01_second_run_20260930.txt)에서 `1 passed in 2.13s`를 확인했습니다. URL과 상품 표시 두 조건의 의미는 선택형으로 확인했고, 입력값·선택자 선택 이유를 본인 말로 설명하는 단계는 남아 있습니다. TC04·TC05·TC02·TC03 코드는 미작성입니다.
- 계획: TC별 수행 절차·UI 검증은 Python에 작성하고 TC-ID에 연결된 입력값·비교값은 엑셀에서 읽습니다. TC01 데이터 한 행은 준비됐지만 현재 로그인 테스트는 코드에 적힌 고정값을 사용합니다. 엑셀 데이터와 로그인 테스트의 연결, 새 버전 엑셀에 자동 결과를 저장하는 기능, 실패 분석 실습은 아직 미구현·미수행입니다.

## 첫 수동 확인

1. https://www.saucedemo.com/ 에 접속합니다.
2. 공개 데모 계정 `standard_user` / `secret_sauce`로 로그인합니다. 페이지에 공개된 계정 정보가 달라졌으면 현재 페이지를 확인합니다.
3. 로그인 전후 화면에서 달라진 점을 보고, 성공 판단 기준을 본인 말로 한 문장 적습니다.
4. 접속이나 로그인이 안 되면 관찰한 오류를 전달합니다. 확인되지 않은 내용을 Pass로 적지 않습니다.

변수·함수부터 짧게 복습하며 TC01 코드를 공동 작성합니다. 다음 테스트는 앞의 예제를 응용해 사용자가 작성하고 AI가 설명·힌트·검토를 제공합니다.

## 실행 환경

AI가 Codex에 포함된 Python 3.12.14를 바탕으로 이 폴더 안의 `.venv`를 준비했습니다. 시스템 Python이나 다른 프로젝트의 패키지는 변경하지 않았습니다. 이 가상환경은 현재 PC의 Codex 런타임에 의존하므로 다른 PC로 폴더를 옮길 때는 Python을 준비한 뒤 재생성합니다.

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

TC01 개별 실행 명령은 다음과 같습니다. 전체 실행·증거 저장 안내는 나머지 테스트 작성에 맞춰 추가합니다.

```powershell
.\.venv\Scripts\python.exe -m pytest .\tests\test_tc01_login.py -q --headed --tracing=retain-on-failure
```

## 기여 구분

- 사용자: 성공 기준 제시, TC01 코드 직접 입력·수정·응용, 수동 확인과 자동화 실행, TC 결과 정정, 엑셀 데이터 작성, 엑셀 읽기 연습 코드 작성·실행.
- AI: 현재 학습 단계의 코드 예시·설명·검토, TC 초안, 환경 설치·점검, 문서 정리, Git 초기 설정을 지원했습니다. TC01과 엑셀 읽기는 예시를 참고해 함께 진행한 학습 작업입니다.
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

## 참고

- [Playwright Python 시작하기](https://playwright.dev/python/docs/intro)
- [테스트 작성](https://playwright.dev/python/docs/writing-tests)
- [요소 선택](https://playwright.dev/python/docs/locators)
- [Trace Viewer](https://playwright.dev/python/docs/trace-viewer-intro)

기존 Selenium 업무 경험과 이번 학습 프로젝트는 별도 사례입니다. VMS·모바일 검증, 성능 개선 수치, 사용자 직접 수행 범위를 추정하지 않습니다.
