from playwright.sync_api import expect

def test_login_success(page):
    username = "standard_user"
    password = "secret_sauce"

    page.goto("https://www.saucedemo.com/")
    page.get_by_placeholder("Username").fill(username)   # Username 입력칸에 username 입력
    page.get_by_placeholder("Password").fill(password)   # Password 입력칸에 password 입력

    page.get_by_role("button", name="Login").click()      # Login 버튼 클릭

    expect(page).to_have_url("https://www.saucedemo.com/inventory.html")  # 로그인 성공 후 URL 확인
    expect(page.locator(".inventory_item").first).to_be_visible()  # 첫 번째 상품이 보이는지 확인