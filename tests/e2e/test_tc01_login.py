"""역할: SauceDemo의 TC01 정상 로그인 시나리오를 실행하는 테스트입니다.

tc_data.py에서 엑셀 TC01의 계정·예상 URL을 받아 Playwright로 로그인하고,
로그인 후 URL과 상품 표시를 확인합니다. pytest가 page 브라우저를 제공합니다.
초기 코드는 사용자 공동 작성, 2026-10-19 엑셀 연결은 AI 작성입니다.
"""

from playwright.sync_api import expect

from tc_data import load_login_data

def test_login_success(page):
    data = load_login_data("TC01")
    username = data["username"]
    password = data["password"]

    page.goto("https://www.saucedemo.com/")
    page.get_by_placeholder("Username").fill(username)   # Username 입력칸에 username 입력
    page.get_by_placeholder("Password").fill(password)   # Password 입력칸에 password 입력

    page.get_by_role("button", name="Login").click()      # Login 버튼 클릭

    expect(page).to_have_url(data["expected_url"])  # 엑셀의 예상 URL과 비교
    expect(page.locator(".inventory_item").first).to_be_visible()  # 첫 번째 상품이 보이는지 확인
