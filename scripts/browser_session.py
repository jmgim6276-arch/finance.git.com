#!/usr/bin/env python3
import getpass
import hashlib
import json
import os
import subprocess
import time
from pathlib import Path
from urllib.parse import quote

import requests
import websocket

BASE_URL = "https://cst.uf-tree.com"
LOGIN_URL = f"{BASE_URL}/login"

BROWSERS = [
    {
        "name": "Edge",
        "port": 9223,
        "url": "http://localhost:9223/json",
        "binary": "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "profile_dir": str(Path.home() / ".finance-cst" / "edge-cdp-profile"),
    },
    {
        "name": "Chrome",
        "port": 18800,
        "url": "http://localhost:18800/json",
        "binary": "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "profile_dir": str(Path.home() / ".finance-cst" / "chrome-cdp-profile"),
    },
]


def is_ok(resp):
    return resp.get("code") == 200 or resp.get("success") is True


def sha1_hex(value):
    return hashlib.sha1(value.encode("utf-8")).hexdigest()


def browser_choices(preferred="auto"):
    preferred = (preferred or "auto").lower()
    if preferred == "edge":
        return [b for b in BROWSERS if b["name"] == "Edge"]
    if preferred == "chrome":
        return [b for b in BROWSERS if b["name"] == "Chrome"]
    return BROWSERS


def list_pages(browser):
    return requests.get(browser["url"], timeout=6).json()


def find_browser(preferred="auto", require_cst=False):
    available = []
    for browser in browser_choices(preferred):
        try:
            pages = list_pages(browser)
            has_cst = any("cst.uf-tree.com" in p.get("url", "") for p in pages)
            if not require_cst or has_cst:
                available.append({**browser, "has_cst": has_cst})
        except Exception:
            continue

    if not available:
        return None

    for browser in available:
        if browser.get("has_cst"):
            return browser
    return available[0]


def wait_for_browser(browser, timeout=20):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            list_pages(browser)
            return True
        except Exception:
            time.sleep(1)
    return False


def launch_browser(preferred="auto", target_url=LOGIN_URL):
    for browser in browser_choices(preferred):
        binary = Path(browser["binary"])
        if not binary.exists():
            continue

        profile_dir = Path(browser["profile_dir"])
        profile_dir.mkdir(parents=True, exist_ok=True)
        cmd = [
            str(binary),
            f"--remote-debugging-port={browser['port']}",
            "--remote-allow-origins=*",
            f"--user-data-dir={profile_dir}",
            target_url,
        ]
        subprocess.Popen(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
        if wait_for_browser(browser):
            return {**browser, "has_cst": True}
    return None


def find_or_launch_browser(preferred="auto", target_url=LOGIN_URL):
    browser = find_browser(preferred=preferred, require_cst=False)
    if browser:
        return browser
    return launch_browser(preferred=preferred, target_url=target_url)


def open_target(browser, url):
    encoded = quote(url, safe=":/?&=#")
    resp = requests.put(
        f"http://localhost:{browser['port']}/json/new?{encoded}",
        timeout=10,
    )
    return resp.json()


def get_cst_page(browser):
    pages = list_pages(browser)
    for page in pages:
        if "cst.uf-tree.com" in page.get("url", ""):
            return page
    return None


def ensure_cst_page(browser, url=LOGIN_URL):
    page = get_cst_page(browser)
    if page:
        return page
    open_target(browser, url)
    deadline = time.time() + 15
    while time.time() < deadline:
        page = get_cst_page(browser)
        if page:
            return page
        time.sleep(1)
    raise RuntimeError(f"{browser['name']} 未能打开财税通页面")


def cdp_eval(page, expression, return_by_value=True):
    ws = websocket.create_connection(
        page["webSocketDebuggerUrl"],
        timeout=10,
        suppress_origin=True,
    )
    try:
        ws.send(
            json.dumps(
                {
                    "id": 1,
                    "method": "Runtime.evaluate",
                    "params": {
                        "expression": expression,
                        "returnByValue": return_by_value,
                    },
                }
            )
        )
        for _ in range(20):
            msg = json.loads(ws.recv())
            if msg.get("id") != 1:
                continue
            result = msg.get("result", {}).get("result", {})
            return result.get("value") if return_by_value else result
    finally:
        ws.close()
    return None


def cdp_navigate(page, url):
    ws = websocket.create_connection(
        page["webSocketDebuggerUrl"],
        timeout=10,
        suppress_origin=True,
    )
    try:
        ws.send(
            json.dumps(
                {
                    "id": 1,
                    "method": "Page.navigate",
                    "params": {"url": url},
                }
            )
        )
        for _ in range(20):
            msg = json.loads(ws.recv())
            if msg.get("id") == 1:
                return msg
    finally:
        ws.close()
    return None


def get_vuex_raw(page):
    return cdp_eval(page, "localStorage.getItem('vuex')")


def parse_vuex(raw_value):
    if not raw_value:
        return {}
    try:
        return json.loads(raw_value)
    except Exception:
        return {}


def extract_auth(page):
    data = parse_vuex(get_vuex_raw(page))
    user = data.get("user", {})
    token = user.get("token")
    company = user.get("company") or {}
    company_id = company.get("id")
    user_id = user.get("id")
    return token, company_id, user_id, data


def validate_auth(token, company_id):
    if not token or not company_id:
        return False
    headers = {"x-token": token, "Content-Type": "application/json"}
    try:
        resp = requests.post(
            f"{BASE_URL}/api/member/department/queryCompany",
            headers=headers,
            json={"companyId": company_id},
            timeout=12,
        ).json()
        return is_ok(resp)
    except Exception:
        return False


def normalize_company_id(company_id):
    if company_id in (None, "", 0):
        return None
    return int(company_id)


def read_credentials(username=None, password=None, company_id=None, prompt=False):
    username = username or os.getenv("CST_USERNAME")
    password = password or os.getenv("CST_PASSWORD")
    company_id = company_id or os.getenv("CST_COMPANY_ID")
    if company_id not in (None, "", 0):
        company_id = int(company_id)

    if prompt and not username:
        username = input("财税通手机号: ").strip()
    if prompt and not password:
        password = getpass.getpass("财税通密码: ").strip()
    return username, password, company_id


def wait_for(condition, timeout=20, interval=1):
    deadline = time.time() + timeout
    while time.time() < deadline:
        value = condition()
        if value:
            return value
        time.sleep(interval)
    return None


def wait_for_login_form(page, timeout=20):
    def _ready():
        state = cdp_eval(
            page,
            """
            JSON.stringify({
              url: location.href,
              ready: [...document.querySelectorAll('input')].some(i => (i.placeholder || '').includes('手机号'))
                && [...document.querySelectorAll('input')].some(i => (i.placeholder || '').includes('密码'))
                && [...document.querySelectorAll('button')].some(b => (b.innerText || '').includes('登录'))
            })
            """,
        )
        try:
            info = json.loads(state or "{}")
        except Exception:
            return None
        return info if info.get("ready") else None

    return wait_for(_ready, timeout=timeout, interval=1)


def submit_login(page, username, password):
    payload = {
        "username": username,
        "password": password,
    }
    script = f"""
    (() => {{
      function setNativeValue(el, value) {{
        const proto = Object.getPrototypeOf(el);
        const descriptor = Object.getOwnPropertyDescriptor(proto, 'value');
        if (descriptor && descriptor.set) {{
          descriptor.set.call(el, value);
        }} else {{
          el.value = value;
        }}
        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
      }}

      const payload = {json.dumps(payload, ensure_ascii=False)};
      const pwdTab = [...document.querySelectorAll('.el-tabs__item')]
        .find(el => (el.innerText || '').includes('密码'));
      if (pwdTab && !pwdTab.classList.contains('is-active')) {{
        pwdTab.click();
      }}
      const usernameInput = [...document.querySelectorAll('input')]
        .find(i => (i.placeholder || '').includes('手机号'));
      const passwordInput = [...document.querySelectorAll('input')]
        .find(i => (i.placeholder || '').includes('密码'));
      const button = [...document.querySelectorAll('button')]
        .find(b => (b.innerText || '').includes('登录'));
      if (!usernameInput || !passwordInput || !button) {{
        return JSON.stringify({{ ok: false, reason: 'login-form-not-found' }});
      }}
      setNativeValue(usernameInput, payload.username);
      setNativeValue(passwordInput, payload.password);
      button.click();
      return JSON.stringify({{ ok: true }});
    }})()
    """
    result = cdp_eval(page, script)
    return json.loads(result or "{}")


def query_user_companies(token):
    headers = {"x-token": token, "Content-Type": "application/json"}
    resp = requests.post(
        f"{BASE_URL}/api/member/userCompanyInfo/queryUserCompany",
        headers=headers,
        json={},
        timeout=12,
    ).json()
    if not is_ok(resp):
        return []
    return resp.get("result") or []


def choose_company(companies, desired_company_id=None):
    desired_company_id = normalize_company_id(desired_company_id)
    if desired_company_id:
        for company in companies:
            if int(company.get("id")) == desired_company_id:
                return company
        raise RuntimeError(f"未在企业列表中找到 companyId={desired_company_id}")
    if len(companies) == 1:
        return companies[0]
    raise RuntimeError("检测到多个企业，请通过 --company-id 或 CST_COMPANY_ID 指定导入企业")


def click_company_entry(page, company_name):
    script = f"""
    (() => {{
      const targetName = {json.dumps(company_name, ensure_ascii=False)};
      const companyCard = [...document.querySelectorAll('.comp')]
        .find(card => (card.innerText || '').includes(targetName));
      if (!companyCard) {{
        return JSON.stringify({{ ok: false, reason: 'company-card-not-found' }});
      }}
      const button = [...companyCard.querySelectorAll('button')]
        .find(btn => (btn.innerText || '').includes('进入企业'));
      if (!button) {{
        return JSON.stringify({{ ok: false, reason: 'enter-button-not-found' }});
      }}
      button.click();
      return JSON.stringify({{ ok: true }});
    }})()
    """
    result = cdp_eval(page, script)
    return json.loads(result or "{}")


def ensure_company_selected(page, desired_company_id=None):
    token, current_company_id, _, _ = extract_auth(page)
    if current_company_id:
        return token, current_company_id

    companies = query_user_companies(token)
    company = choose_company(companies, desired_company_id=desired_company_id)
    click = click_company_entry(page, company["name"])
    if not click.get("ok"):
        raise RuntimeError(f"进入企业失败：{click}")

    def _selected():
        token2, company_id2, _, _ = extract_auth(page)
        if token2 and company_id2 and int(company_id2) == int(company["id"]):
            return token2, int(company_id2)
        return None

    selected = wait_for(_selected, timeout=20, interval=1)
    if not selected:
        raise RuntimeError("企业选择后未能读取到有效 companyId")
    return selected


def ensure_login(
    preferred_browser="auto",
    username=None,
    password=None,
    company_id=None,
    prompt=False,
):
    browser = find_or_launch_browser(preferred=preferred_browser, target_url=LOGIN_URL)
    if not browser:
        raise RuntimeError("未能自动打开 Edge/Chrome，请确认本机已安装浏览器")

    page = ensure_cst_page(browser, url=LOGIN_URL)
    token, selected_company_id, user_id, _ = extract_auth(page)
    if validate_auth(token, selected_company_id):
        return token, selected_company_id, user_id, browser["name"]

    username, password, company_id = read_credentials(
        username=username,
        password=password,
        company_id=company_id,
        prompt=prompt,
    )
    if not username or not password:
        raise RuntimeError(
            "当前登录态失效，且未提供账号密码。可加 --auto-login，并通过 --username + 终端隐藏输入，或设置 CST_USERNAME/CST_PASSWORD。"
        )

    cdp_navigate(page, LOGIN_URL)
    if not wait_for_login_form(page, timeout=20):
        raise RuntimeError("登录页未就绪，无法自动填写账号密码")

    submit = submit_login(page, username=username, password=password)
    if not submit.get("ok"):
        raise RuntimeError(f"自动提交登录失败：{submit}")

    def _login_ready():
        token2, company_id2, user_id2, _ = extract_auth(page)
        if token2:
            return token2, company_id2, user_id2
        return None

    logged_in = wait_for(_login_ready, timeout=20, interval=1)
    if not logged_in:
        raise RuntimeError("自动登录后未能读取到 token，请检查账号密码是否正确")

    token, selected_company_id, user_id = logged_in
    if not selected_company_id:
        token, selected_company_id = ensure_company_selected(page, desired_company_id=company_id)

    if not validate_auth(token, selected_company_id):
        raise RuntimeError("自动登录完成，但登录态校验未通过")

    return token, selected_company_id, user_id, browser["name"]


def get_auth(
    auto_login=False,
    preferred_browser="auto",
    username=None,
    password=None,
    company_id=None,
    prompt=False,
):
    browser = find_or_launch_browser(preferred=preferred_browser, target_url=LOGIN_URL)
    if not browser:
        raise RuntimeError(
            "未检测到可用的浏览器。请安装 Edge 或 Chrome，或使用 --auto-login 让脚本自动打开浏览器。"
        )

    page = ensure_cst_page(browser, url=LOGIN_URL)
    token, selected_company_id, user_id, _ = extract_auth(page)
    if validate_auth(token, selected_company_id):
        return token, selected_company_id, user_id, browser["name"]

    if not auto_login:
        raise RuntimeError(
            "浏览器中的财税通登录态已失效。请重新登录，或使用 --auto-login 让脚本自动登录。"
        )

    return ensure_login(
        preferred_browser=preferred_browser,
        username=username,
        password=password,
        company_id=company_id,
        prompt=prompt,
    )
