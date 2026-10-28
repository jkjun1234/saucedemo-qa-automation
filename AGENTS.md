# 엑셀 서식 비교 시 주의

2026-10-28 사용자 요청으로 기록한 프로젝트 한정 팁입니다.

- openpyxl의 `cell.font`는 서식 정보를 감싼 `StyleProxy`로 제공됩니다. 두 셀의 폰트가 같아도 이 객체끼리 직접 비교하면 불일치로 판정할 수 있습니다.
- 서식 보존 검사에서는 `copy()`로 실제 폰트 설정값을 복사해 비교합니다.

```python
from copy import copy

assert copy(prepared_cell.font) == copy(original_cell.font)
```

- 여기서 복사하는 것은 메모리의 폰트 설정이며 엑셀 파일 사본을 생성하는 것이 아닙니다.
- 이 비교 코드의 오류를 실제 서비스 결함이나 사용자 테스트 수행 실적으로 기록하지 않습니다.

근거: [openpyxl StyleProxy 공식 구현](https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/proxy.html)
