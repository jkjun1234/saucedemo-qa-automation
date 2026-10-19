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
