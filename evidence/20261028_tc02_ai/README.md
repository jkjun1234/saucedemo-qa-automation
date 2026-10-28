# TC02 잠긴 계정 로그인 차단 — AI 검증

2026-10-28, 작성·관찰·자동화 실행·엑셀 반영은 AI(Codex)가 수행했습니다. 사용자 직접 수행 실적으로 계산하지 않습니다.

## 기대 결과의 근거

실제 SauceDemo에서 공개된 `locked_out_user` / `secret_sauce`를 입력한 결과:

- `Epic sadface: Sorry, this user has been locked out.` 표시.
- 주소 `https://www.saucedemo.com/` 유지.
- Login 버튼 표시, `.inventory_item` 0개.

`observation.json`, `locked-account-observation.png`는 구현 전 AI의 실제 브라우저 관찰입니다. 서비스 결함 발견 사례가 아닙니다.

## 자동화와 결과

- 실행용 엑셀 TC02: 접속 → 계정·비밀번호 입력 → Login 클릭 → 오류 문구 일치 → 로그인 URL 유지 → Login 버튼 표시 → 상품 카드 정확히0개. 총8단계.
- 오류가 기대대로 발생하므로 업무 TC 결과는 **Pass**입니다. 단순히 오류가 발생했다는 이유로 Fail을 기록하지 않습니다.
- 전체 검사 **55 passed in21.89s**: 개발52개 + 실제 사이트 검사3개(TC01 기존/새 방식·TC02). 업무 시나리오는2개이며55개 시나리오 완성이 아닙니다. `pytest.log`, `pytest.xml`에 보관했습니다.
- 별도 보이는 Chromium 실행 `20261028_222028_c91b9f01`: 8단계 완료·TC02 Pass·선택한 TC 파일 자체 저장 성공. `tc02-events.jsonl`, `tc02-summary.json`을 확인합니다. 이 실행은 GUI의 시작 버튼이 아닌 같은 실행기와 결과 저장 함수를 호출한 통합 검증입니다.
- `문구확인`·`개수확인`은 실제 로컬 페이지에서 맞는 값은 Pass, 틀린 값은 Fail이 되는지 검사했습니다. 결과를 모의 생성하거나 기대값에 맞춰 실제 화면을 바꾸지 않았습니다.

## 파일 보존과 복구

- TC02 준비 직전 사용자가 저장한 실행용 파일을 기준으로 기존 셀 값·서식을 비교했습니다. 기존 셀은 안내 A1과 TC02 결과 H4만 변경, 실행 단계8행 추가. TC01·TC03~05·기존 설명과 서식은 유지했습니다(`workbook-check.json`).
- 이전 설계서 초안 원본의 바이트 해시는 Git 원본과 같았습니다.
- 사용자가 A를 선택해 실행기 파일을 검증된 `45ea31f`로 복구한 후 TC02를 추가했습니다. 복구 전 파일은 Git에서 제외된 로컬 작업 기록에 보존했습니다. 중지·환경 오류가 Fail로 잘못 기록되던 기존 검사2개는 복구 전 실패를 재현했고 최종 전체 검사에서 통과했습니다. 변경 주체를 추측하지 않습니다.
- 사용자 직접 TC02 실행·결과 해석은 다음 확인 단계입니다. 현재 창은 새 코드를 받도록 닫고 다시 열어야 합니다.

참고: [Playwright 문구·개수 검증 API](https://playwright.dev/python/docs/api/class-locatorassertions).
