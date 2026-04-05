#!/usr/bin/env python3
import argparse

import requests

from browser_session import BASE_URL, get_auth


def check_get(url, headers, params, name):
    r = requests.get(url, headers=headers, params=params, timeout=12)
    j = r.json()
    ok = j.get("code") == 200 or j.get("success") is True
    print(("✅" if ok else "❌"), name)
    return ok


def check_post(url, headers, payload, name):
    r = requests.post(url, headers=headers, json=payload, timeout=12)
    j = r.json()
    ok = j.get("code") == 200 or j.get("success") is True
    print(("✅" if ok else "❌"), name)
    return ok


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="财税通环境预检查")
    parser.add_argument("--auto-login", action="store_true", help="登录态失效时自动打开浏览器并登录")
    parser.add_argument("--username", help="财税通登录手机号；不传则优先读取 CST_USERNAME，仍缺失时终端提示输入")
    parser.add_argument("--company-id", type=int, help="多企业账号时指定 companyId；也可用环境变量 CST_COMPANY_ID")
    parser.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome"],
        default="auto",
        help="优先使用的浏览器",
    )
    args = parser.parse_args()

    token, company_id, _, browser_name = get_auth(
        auto_login=args.auto_login,
        preferred_browser=args.browser,
        username=args.username,
        company_id=args.company_id,
        prompt=args.auto_login,
    )
    headers = {"x-token": token, "Content-Type": "application/json"}
    print(f"✅ 检测到 {browser_name} 浏览器")
    print(f"✅ 登录态可读: companyId={company_id}")

    checks = [
        check_post(f"{BASE_URL}/api/member/department/queryCompany", headers, {"companyId": company_id}, "queryCompany"),
        check_get(f"{BASE_URL}/api/member/department/queryDepartments", headers, {"companyId": company_id}, "queryDepartments"),
        check_get(f"{BASE_URL}/api/member/role/get/tree", headers, {"companyId": company_id}, "role/get/tree"),
        check_get(f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate", headers, {"companyId": company_id, "status": 1, "pageSize": 1000}, "queryFeeTemplate"),
        check_get(f"{BASE_URL}/api/bpm/workflow/queryWorkFlow", headers, {"companyId": company_id, "size": 200}, "queryWorkFlow"),
        check_get(f"{BASE_URL}/api/bill/template/queryTemplateTree", headers, {"companyId": company_id}, "queryTemplateTree"),
    ]

    if all(checks):
        print("\n✅ PRECHECK PASS")
    else:
        print("\n❌ PRECHECK FAIL")
        raise SystemExit(1)
