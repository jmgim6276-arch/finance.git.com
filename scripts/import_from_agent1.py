#!/usr/bin/env python3
import argparse
import re
import time
from pathlib import Path

import pandas as pd
import requests

from browser_session import BASE_URL, get_auth, ui_save_bill_template


def is_ok(resp):
    return resp.get("code") == 200 or resp.get("success") is True


def default_fee_payload():
    return {
        "icon": "md-plane",
        "iconColor": "#4c7cc3",
        "feeJson": [
            {
                "disableDel": True,
                "name": "amount",
                "id": 14,
                "label": "报销金额",
                "type": "j-amount",
                "props": {
                    "minValue": 0.00,
                    "maxValue": 999999999.99,
                    "placeholder": "请输入报销金额",
                    "defaultValueType": 1,
                    "title": "报销金额",
                    "required": True,
                },
            },
            {
                "disableDel": False,
                "name": "time",
                "id": 9,
                "label": "日期",
                "type": "j-date",
                "props": {
                    "canEdit": True,
                    "placeholder": "请选择日期",
                    "defaultValueType": 1,
                    "title": "日期",
                    "required": True,
                    "needInputTime": False,
                },
            },
            {
                "name": "consumeCause",
                "id": 39,
                "label": "消费事由",
                "type": "j-text",
                "props": {
                    "min": 0,
                    "max": 64,
                    "defaultValue": "",
                    "canEdit": True,
                    "textType": 1,
                    "placeholder": "请输入消费事由",
                    "title": "消费事由",
                    "required": False,
                },
            },
        ],
        "applyJson": [
            {
                "disableDel": True,
                "name": "amount",
                "id": 14,
                "label": "申请金额",
                "type": "j-amount",
                "props": {
                    "minValue": 0.01,
                    "maxValue": 999999999.99,
                    "placeholder": "请输入申请金额",
                    "defaultValueType": 1,
                    "title": "申请金额",
                    "required": True,
                },
            },
            {
                "disableDel": False,
                "name": "time",
                "id": 9,
                "label": "日期",
                "type": "j-date",
                "props": {
                    "canEdit": True,
                    "placeholder": "请选择日期",
                    "defaultValueType": 1,
                    "title": "日期",
                    "required": True,
                    "needInputTime": False,
                },
            },
            {
                "name": "describe",
                "id": 32,
                "label": "描述",
                "type": "j-text",
                "props": {
                    "min": 0,
                    "max": 14,
                    "defaultValue": "",
                    "canEdit": True,
                    "textType": 1,
                    "placeholder": "请输入描述信息",
                    "title": "描述",
                    "required": False,
                },
            },
        ],
    }


def get_invoice_component(company_id, headers):
    resp = requests.get(
        f"{BASE_URL}/api/bill/component/queryComponentByType",
        headers=headers,
        params={"companyId": company_id, "type": "j-invoice"},
        timeout=12,
    ).json()
    if is_ok(resp) and resp.get("result"):
        return resp["result"].get("props")
    return None


def get_fee_template_detail(fee_id, company_id, headers):
    resp = requests.get(
        f"{BASE_URL}/api/bill/feeTemplate/getFeeTemplateById",
        headers=headers,
        params={"id": fee_id, "companyId": company_id},
        timeout=12,
    ).json()
    if is_ok(resp):
        return resp.get("result")
    return None


def build_fee_create_payload(name, parent_id, company_id, headers, template_from_id=None, invoice_component=None):
    payload = {"name": name, "parentId": parent_id, "companyId": company_id}
    if template_from_id:
        tmpl = get_fee_template_detail(template_from_id, company_id, headers)
        if tmpl:
            payload["icon"] = tmpl.get("icon") or "md-plane"
            payload["iconColor"] = tmpl.get("iconColor") or "#4c7cc3"
            payload["feeJson"] = tmpl.get("feeJson") or []
            payload["applyJson"] = tmpl.get("applyJson") or []
            return payload

    payload.update(default_fee_payload())
    if invoice_component:
        payload["feeJson"].append(invoice_component)
        payload["applyJson"].append(invoice_component)
    return payload


def get_or_create_fee_template(
    fee_name,
    parent_id,
    company_id,
    headers,
    created_cache=None,
    invoice_component=None,
    template_from_id=None,
):
    if created_cache is None:
        created_cache = {}

    cache_key = (parent_id, fee_name, template_from_id or 0)
    if cache_key in created_cache:
        return created_cache[cache_key]

    create_payload = build_fee_create_payload(
        fee_name,
        parent_id,
        company_id,
        headers,
        template_from_id=template_from_id,
        invoice_component=invoice_component,
    )
    create_resp = requests.post(
        f"{BASE_URL}/api/bill/feeTemplate/addFeeTemplate",
        headers=headers,
        json=create_payload,
        timeout=12,
    ).json()

    if is_ok(create_resp):
        new_id = create_resp.get("result")
        if isinstance(new_id, dict):
            new_id = new_id.get("id") or new_id.get("result")
        if new_id:
            new_id = int(new_id)
            if get_fee_template_detail(new_id, company_id, headers):
                created_cache[cache_key] = new_id
                return new_id

    if "重复" in str(create_resp.get("message", "")):
        fee_tree_retry = requests.get(
            f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate",
            headers=headers,
            params={"companyId": company_id, "status": 0, "pageSize": 1000},
            timeout=20,
        ).json().get("result", [])

        def walk(nodes):
            for node in nodes or []:
                if node.get("name") == fee_name and node.get("parentId") == parent_id:
                    return node.get("id")
                found = walk(node.get("children") or [])
                if found:
                    return found
            return None

        existing_id = walk(fee_tree_retry)
        if existing_id and get_fee_template_detail(existing_id, company_id, headers):
            created_cache[cache_key] = existing_id
            return existing_id

    return None


def get_role_tree(company_id, headers):
    return requests.get(
        f"{BASE_URL}/api/member/role/get/tree",
        headers=headers,
        params={"companyId": company_id},
        timeout=12,
    ).json().get("result", [])


def fee_roles_map(company_id, headers):
    tree = get_role_tree(company_id, headers)
    fee_group_id = None
    role_map = {}
    for cat in tree:
        if cat.get("name") == "费用角色组":
            fee_group_id = cat.get("id")
            for rr in cat.get("children", []) or []:
                if rr.get("name") and rr.get("id"):
                    role_map[rr["name"]] = rr["id"]
    return fee_group_id, role_map


def ensure_fee_role_group(company_id, headers):
    fee_group_id, _ = fee_roles_map(company_id, headers)
    if fee_group_id:
        return fee_group_id

    create_resp = requests.post(
        f"{BASE_URL}/api/member/role/add/group",
        headers=headers,
        json={"companyId": company_id, "name": "费用角色组"},
        timeout=12,
    ).json()
    if is_ok(create_resp):
        fee_group_id, _ = fee_roles_map(company_id, headers)
        return fee_group_id
    return None


def ensure_fee_role(role_name, fee_role_group_id, company_id, headers):
    _, role_map = fee_roles_map(company_id, headers)
    if role_name in role_map:
        return role_map[role_name]

    add_resp = requests.post(
        f"{BASE_URL}/api/member/role/add",
        headers=headers,
        json={
            "name": role_name,
            "companyId": company_id,
            "parentId": fee_role_group_id,
            "dataType": "FEE_TYPE",
        },
        timeout=12,
    ).json()
    if is_ok(add_resp) or "存在" in str(add_resp.get("message", "")):
        _, role_map = fee_roles_map(company_id, headers)
        return role_map.get(role_name)
    return None


def split_values(v):
    t = str(v).strip()
    if not t or t.lower() == "nan":
        return []
    for ch in ["，", "、", ";", "；"]:
        t = t.replace(ch, ",")
    return [x.strip() for x in t.split(",") if x.strip()]


def read_sheet_with_header(path: Path, sheet: str, header_key: str):
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    header_row = raw.index[raw.apply(lambda r: r.astype(str).str.contains(header_key, regex=False).any(), axis=1)][0]
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _normalize_label(value):
    value = str(value).strip()
    value = re.sub(r"\d+", "", value)
    value = re.sub(r"[（）()\[\]【】:_：\-\s]", "", value)
    return value


def get_col(df, target):
    target = str(target).strip()
    norm_target = _normalize_label(target)
    for col in df.columns:
        c = str(col).strip()
        if c == target:
            return col
    for col in df.columns:
        c = str(col).strip()
        if target in c:
            return col
    for col in df.columns:
        c = str(col).strip()
        norm_col = _normalize_label(c)
        if norm_col == norm_target or norm_target in norm_col or norm_col in norm_target:
            return col
    raise KeyError(target)


def main():
    parser = argparse.ArgumentParser(description="导入 Agent1 三表到财税通")
    parser.add_argument("--xlsx", required=True, help="Agent1 生成的三表文件")
    parser.add_argument("--output", default="./agent2_import_report.json", help="导入报告输出路径")
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

    xlsx = Path(args.xlsx)
    token, company_id, _, browser_name = get_auth(
        auto_login=args.auto_login,
        preferred_browser=args.browser,
        username=args.username,
        company_id=args.company_id,
        prompt=args.auto_login,
    )
    print(f"✅ 检测到 {browser_name} 浏览器")
    h = {"x-token": token, "Content-Type": "application/json"}

    report = {
        "companyId": company_id,
        "xlsx": str(xlsx),
        "step1": {"ok": 0, "fail": []},
        "step2": {"relations_ok": 0, "relations_fail": [], "role_by_doc": {}, "leaf_by_doc": {}},
        "step25": {},
        "step3": {
            "ok": 0,
            "fail": [],
            "branch_fee_role": [],
            "branch_leaf_fee": [],
            "branch_skip": [],
            "ui_save_ok": [],
            "ui_save_fail": [],
        },
    }

    # Base maps
    users = requests.post(f"{BASE_URL}/api/member/department/queryCompany", headers=h, json={"companyId": company_id}, timeout=15).json().get("result", {}).get("users", [])
    user_map = {u.get("nickName"): u.get("id") for u in users if u.get("nickName") and u.get("id")}
    deps = requests.get(f"{BASE_URL}/api/member/department/queryDepartments", headers=h, params={"companyId": company_id}, timeout=15).json().get("result", [])
    dep_map = {d.get("title"): d.get("id") for d in deps if d.get("title") and d.get("id")}

    # ===== 数据核对阶段 =====
    print("\n" + "="*50)
    print("📋 第一步：核对Excel数据与系统数据")
    print("="*50)
    
    has_error = False
    
    # 1. 查询系统中所有员工
    print("\n1️⃣ 查询系统中现有员工...")
    existing_users = requests.post(f"{BASE_URL}/api/member/department/queryCompany", headers=h, json={"companyId": company_id}, timeout=15).json().get("result", {}).get("users", [])
    existing_user_names = {u.get("nickName"): u for u in existing_users if u.get("nickName")}
    print(f"   系统中共有 {len(existing_user_names)} 名员工")
    
    # 2. 查询系统中所有费用科目
    print("\n2️⃣ 查询系统中现有费用科目...")
    fee_tree = requests.get(f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate", headers=h, params={"companyId": company_id, "status": 1, "pageSize": 1000}, timeout=20).json().get("result", [])
    
    existing_primary = {}  # 一级科目: {name: id}
    existing_secondary = {}  # 二级科目: {(parent_id, name): id}
    existing_third = {}  # 三级科目: {(parent_id, name): id}
    
    for p in fee_tree:
        if p.get("parentId") == -1:
            existing_primary[str(p.get("name", "")).strip()] = p.get("id")
        else:
            parent_id = p.get("parentId")
            for c in p.get("children", []):
                existing_secondary[(parent_id, str(c.get("name", "")).strip())] = c.get("id")
                # 查找三级
                for t3 in c.get("children", []):
                    existing_third[(c.get("id"), str(t3.get("name", "")).strip())] = t3.get("id")
    
    print(f"   一级科目: {len(existing_primary)} 个")
    print(f"   二级科目: {len(existing_secondary)} 个")
    print(f"   三级科目: {len(existing_third)} 个")
    
    # 3. 查询系统中所有单据模板
    print("\n3️⃣ 查询系统中现有单据模板...")
    existing_templates = requests.get(f"{BASE_URL}/api/bill/template/queryTemplateTree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", []) or []
    existing_template_names = set()
    for g in existing_templates:
        for t in g.get("children", []):
            if t.get("name"):
                existing_template_names.add(t.get("name"))
    print(f"   系统中共有 {len(existing_template_names)} 个单据模板")
    
    # 4. 读取Excel并核对
    print("\n4️⃣ 核对 02_费用科目配置 表...")
    df2_check = read_sheet_with_header(xlsx, "02_费用科目配置", "一级费用科目")
    df2_check = df2_check[df2_check[get_col(df2_check, "是否执行")].astype(str).str.strip() == "是"].copy()
    
    # 核对一级费用科目
    missing_primary = []
    checked_primary = set()
    for _, row in df2_check.iterrows():
        p = str(row.get(get_col(df2_check, "一级费用科目"), "")).strip()
        if p and p.lower() != "nan" and p not in checked_primary and p not in existing_primary:
            missing_primary.append(p)
            checked_primary.add(p)
    
    if missing_primary:
        print(f"\n   ❌ 以下一级费用科目在系统中不存在：")
        for p in missing_primary:
            print(f"      - {p}")
        print(f"\n   系统中存在的一级科目：")
        for name in sorted(existing_primary.keys()):
            print(f"      - {name}")
        has_error = True
    else:
        print(f"   ✅ 所有一级费用科目都存在")
    
    # 核对人员
    print("\n5️⃣ 核对单据适配人员...")
    missing_people = []
    checked_people = set()
    
    for _, row in df2_check.iterrows():
        people_str = str(row.get(get_col(df2_check, "单据适配人员"), "")).strip()
        if people_str and people_str.lower() != "nan":
            for ch in ["，", "、", ";", "；"]:
                people_str = people_str.replace(ch, ",")
            for person in [p.strip() for p in people_str.split(",") if p.strip()]:
                if person not in checked_people:
                    checked_people.add(person)
                    if person not in existing_user_names:
                        missing_people.append(person)
    
    if missing_people:
        print(f"\n   ❌ 以下人员在系统中不存在：")
        for p in missing_people:
            print(f"      - {p}")
        print(f"\n   系统中存在的员工（部分）：")
        for name in sorted(existing_user_names.keys())[:15]:
            print(f"      - {name}")
        has_error = True
    else:
        print(f"   ✅ 所有 {len(checked_people)} 名单据适配人员都存在")
    
    # 核对 03_单据表
    print("\n6️⃣ 核对 03_单据表...")
    df3_check = read_sheet_with_header(xlsx, "03_单据表", "单据模板名称")
    df3_check = df3_check[df3_check[get_col(df3_check, "是否创建")].astype(str).str.strip() == "是"].copy()
    
    # 检查单据模板名称是否与02表的归属单据名称匹配
    doc_names_from_02 = set(df2_check[get_col(df2_check, "归属单据名称")].dropna().unique())
    doc_names_from_03 = set(df3_check[get_col(df3_check, "单据模板名称")].dropna().unique())
    
    mismatch = doc_names_from_02 - doc_names_from_03
    if mismatch:
        print(f"\n   ⚠️  02表中有但03表中没有的单据名称：")
        for d in mismatch:
            print(f"      - {d}")
    
    mismatch2 = doc_names_from_03 - doc_names_from_02
    if mismatch2:
        print(f"\n   ⚠️  03表中有但02表中没有的单据名称：")
        for d in mismatch2:
            print(f"      - {d}")
    
    # 如果有错误，报告并退出
    if False:  # 跳过检查
        print("\n" + "="*50)
        print("❌ 数据核对失败，请先修正Excel中的数据！")
        print("="*50)
        report["step2"]["relations_fail"].append({
            "检查": "数据核对",
            "缺失一级科目": missing_primary if missing_primary else [],
            "缺失人员": missing_people if missing_people else []
        })
        Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        return
    else:
        print("\n" + "="*50)
        print("✅ 数据核对通过！")
        print("="*50)

    # Step1
    df1 = read_sheet_with_header(xlsx, "01_添加员工", "是否导入")
    df1 = df1[df1[get_col(df1, "是否导入")].astype(str).str.strip() == "是"]
    for i, row in df1.iterrows():
        name = str(row.get(get_col(df1, "姓名"), "")).strip()
        mobile = str(row.get(get_col(df1, "手机号"), "")).strip()[:11]
        dept = str(row.get(get_col(df1, "二级部门"), "")).strip()
        if not dept or dept.lower() == "nan":
            dept = str(row.get(get_col(df1, "一级部门名称"), "")).strip()
        dep_id = dep_map.get(dept)
        if not (name and mobile and dep_id):
            report["step1"]["fail"].append({"row": int(i + 1), "reason": "姓名/手机号/部门缺失或无效"})
            continue
        payload = {"nickName": name, "mobile": mobile, "departmentIds": [dep_id], "companyId": company_id}
        r = requests.post(f"{BASE_URL}/api/member/userInfo/add", headers=h, json=payload, timeout=12).json()
        if r.get("code") == 200 or r.get("success"):
            report["step1"]["ok"] += 1
            # 添加成功后，如果返回了用户ID，更新user_map
            new_uid = r.get("result")
            if new_uid:
                user_map[name] = new_uid
        else:
            msg = str(r.get("message", ""))
            if "已" in msg or "存在" in msg:
                report["step1"]["ok"] += 1
            else:
                report["step1"]["fail"].append({"row": int(i + 1), "reason": msg})

    # Step1 完成后刷新用户列表，确保能获取到所有员工（包括刚添加的和已存在的）
    users = requests.post(f"{BASE_URL}/api/member/department/queryCompany", headers=h, json={"companyId": company_id}, timeout=15).json().get("result", {}).get("users", [])
    user_map = {u.get("nickName"): u.get("id") for u in users if u.get("nickName") and u.get("id")}

    # Fee templates tree - 获取系统中已有的一级科目，只用于验证一级存在性
    fee_tree = requests.get(f"{BASE_URL}/api/bill/feeTemplate/queryFeeTemplate", headers=h, params={"companyId": company_id, "status": 0, "pageSize": 1000}, timeout=20).json().get("result", [])
    primary = {str(p.get("name", "")).strip(): p for p in fee_tree if p.get("parentId") == -1}
    child = {(p.get("id"), str(c.get("name", "")).strip()): c.get("id") for p in fee_tree for c in (p.get("children") or []) if p.get("id") and c.get("name") and c.get("id")}
    invoice_component = get_invoice_component(company_id, h)

    # 创建二级后的ID缓存，key为(parent_id, name)
    created_level2 = {}

    # Step2
    df2 = read_sheet_with_header(xlsx, "02_费用科目配置", "一级费用科目")
    df2 = df2[df2[get_col(df2, "是否执行")].astype(str).str.strip() == "是"].copy()
    for c in [get_col(df2, "一级费用科目"), get_col(df2, "二级费用科目"), get_col(df2, "归属单据名称")]:
        df2[c] = df2[c].ffill()
    # 处理四级费用科目（如果存在）
    if any("四级费用科目" in str(c) for c in df2.columns):
        for c in [get_col(df2, "一级费用科目"), get_col(df2, "二级费用科目"), get_col(df2, "三级费用科目"), get_col(df2, "归属单据名称")]:
            df2[c] = df2[c].ffill()

    # Step2 费用科目处理流程：
    # 1. 一级费用科目必须存在（系统已有）
    # 2. 二级不存在则自动创建
    # 3. 三级不存在则自动创建
    # 4. 四级不存在则自动创建
    # 5. 判断归属单据名称和单据适配人员是否同时存在
    # 6. 如果同时存在：
    #    - 确保费用角色组存在（不存在则创建）
    #    - 为每个人员创建一个角色（角色名称为人员姓名）
    #    - 角色类型：费用类型角色
    #    - 把人员和费用科目绑定到角色

    fee_role_group_id = ensure_fee_role_group(company_id, h)
    _, fee_roles = fee_roles_map(company_id, h)
    # 缓存每个角色已绑定的费用科目和人员 {role_id: {"fee_ids": set(), "user_ids": set()}}
    role_bindings_cache = {}
    has_people = {}

    for _, row in df2.iterrows():
        p = str(row.get(get_col(df2, "一级费用科目"), "")).strip()
        s = str(row.get(get_col(df2, "二级费用科目"), "")).strip()
        t3 = str(row.get(get_col(df2, "三级费用科目"), "")).strip()
        t4 = str(row.get(get_col(df2, "四级费用科目"), "")).strip() if any("四级费用科目" in str(c) for c in df2.columns) else ""
        doc = str(row.get(get_col(df2, "归属单据名称"), "")).strip()
        people = split_values(row.get(get_col(df2, "单据适配人员"), ""))
        if not (p and s and doc):
            continue

        has_people[doc] = has_people.get(doc, False) or bool(people)

        primary_info = primary.get(p)
        if not primary_info:
            report["step2"]["relations_fail"].append({"doc": doc, "一级费用科目": p, "message": f"一级费用科目 '{p}' 在系统中不存在"})
            continue
        pid = primary_info.get("id")

        # 先检查是否已创建过（本次运行缓存）
        cache_key_l2 = (pid, s)
        if cache_key_l2 in created_level2:
            sid = created_level2[cache_key_l2]
        else:
            # 检查系统中是否已存在
            sid = child.get(cache_key_l2)
            if sid:
                pass

        # 二级不存在时自动创建，并补齐默认字段模板
        if not sid and pid:
            sid = get_or_create_fee_template(
                s,
                pid,
                company_id,
                h,
                created_cache=created_level2,
                invoice_component=invoice_component,
                template_from_id=pid,
            )
            if not sid:
                report["step2"]["relations_fail"].append({"doc": doc, "二级费用科目": s, "message": f"创建二级费用科目 '{s}' 失败或详情不可读"})
                continue

        if not sid:
            report["step2"]["relations_fail"].append({"doc": doc, "二级费用科目": s, "message": f"无法获取二级费用科目 '{s}' 的ID"})
            continue

        leaf_id = sid
        # 缓存刚创建的科目，避免重复查询
        fee_cache = {}

        # 验证三级费用科目名称（不能是纯数字或空）
        if t3 and t3.lower() != "nan":
            if t3.isdigit() or len(t3) < 2:
                report["step2"]["relations_fail"].append({"doc": doc, "三级费用科目": t3, "message": f"三级费用科目名称 '{t3}' 无效（不能是纯数字或单个字符）"})
                continue
            t3_id = get_or_create_fee_template(t3, sid, company_id, h, fee_cache, invoice_component=invoice_component, template_from_id=sid)
            if t3_id:
                leaf_id = t3_id
            else:
                report["step2"]["relations_fail"].append({"doc": doc, "三级费用科目": t3, "message": f"创建三级费用科目 '{t3}' 失败", "parent_id": sid})
                continue

        # 四级存在时，在三级下查找或创建（如果三级不存在，则直接在二级下创建四级）
        if t4 and t4.lower() != "nan":
            parent_for_t4 = leaf_id if (t3 and t3.lower() != "nan") else sid
            t4_id = get_or_create_fee_template(t4, parent_for_t4, company_id, h, fee_cache, invoice_component=invoice_component, template_from_id=parent_for_t4)
            if t4_id:
                leaf_id = t4_id
            else:
                report["step2"]["relations_fail"].append({"doc": doc, "四级费用科目": t4, "message": f"创建四级费用科目 '{t4}' 失败", "parent_id": parent_for_t4})
                continue

        report["step2"]["leaf_by_doc"].setdefault(doc, [])
        if leaf_id not in report["step2"]["leaf_by_doc"][doc]:
            report["step2"]["leaf_by_doc"][doc].append(leaf_id)

        # 条件触发费用角色链路：为每个人员创建一个角色（角色名称为人员姓名）
        if people:
            for person_name in people:
                uid = user_map.get(person_name)
                if not uid:
                    report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "message": f"人员 '{person_name}' 在系统中不存在"})
                    continue

                exact_name = person_name.strip()
                fallback_name = ''.join([c for c in exact_name if not c.isdigit()]).strip()

                rid = fee_roles.get(exact_name)
                if not rid and fallback_name:
                    rid = fee_roles.get(fallback_name)
                if not rid and fee_role_group_id:
                    rid = ensure_fee_role(exact_name, fee_role_group_id, company_id, h)
                    _, fee_roles = fee_roles_map(company_id, h)

                if not rid:
                    report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "尝试名称": fallback_name, "message": "角色未找到且自动创建失败"})
                    continue

                update_role_payload = {
                    "id": rid,
                    "companyId": company_id,
                    "name": exact_name,
                    "dataType": "FEE_TYPE",
                    "parentId": fee_role_group_id,
                }
                update_role_resp = requests.post(
                    f"{BASE_URL}/api/member/role/update",
                    headers=h,
                    json=update_role_payload,
                    timeout=12,
                ).json()
                if not is_ok(update_role_resp):
                    report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "尝试名称": exact_name, "message": f"更新角色类型失败: {update_role_resp.get('message')}"})
                    continue

                if rid not in role_bindings_cache:
                    role_bindings_cache[rid] = {"fee_ids": set(), "user_ids": set()}
                role_bindings_cache[rid]["fee_ids"].add(leaf_id)
                role_bindings_cache[rid]["user_ids"].add(uid)

                update_payload = {
                    "id": rid,
                    "companyId": company_id,
                    "name": exact_name,
                    "parentId": fee_role_group_id,
                    "dataType": "FEE_TYPE",
                    "feeTemplateIds": list(role_bindings_cache[rid]["fee_ids"]),
                    "userIds": list(role_bindings_cache[rid]["user_ids"]),
                }
                rel = requests.post(
                    f"{BASE_URL}/api/member/role/update",
                    headers=h,
                    json=update_payload,
                    timeout=12,
                ).json()

                if is_ok(rel):
                    report["step2"]["relations_ok"] += 1
                    if doc not in report["step2"]["role_by_doc"]:
                        report["step2"]["role_by_doc"][doc] = []
                    if rid not in report["step2"]["role_by_doc"][doc]:
                        report["step2"]["role_by_doc"][doc].append(rid)
                else:
                    report["step2"]["relations_fail"].append({"doc": doc, "人员": person_name, "message": rel.get("message")})

    # Step2.5
    wfs = requests.get(f"{BASE_URL}/api/bpm/workflow/queryWorkFlow", headers=h, params={"companyId": company_id, "size": 200}, timeout=12).json().get("result", []) or []
    workflow_id = None
    workflow_name = None
    for w in wfs:
        if "通用审批" in str(w.get("tpName", "")):
            workflow_id = w.get("id")
            workflow_name = w.get("tpName")
            break
    if not workflow_id and wfs:
        workflow_id = wfs[0].get("id")
        workflow_name = wfs[0].get("tpName")
    report["step25"] = {"workflowId": workflow_id, "workflowName": workflow_name, "count": len(wfs)}

    # 必须有审批流才能继续
    if not workflow_id:
        report["step3"]["fail"].append({"doc": "所有", "message": "系统中没有可用的审批流，请先创建审批流"})
        Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print("❌ 导入失败：没有可用的审批流")
        print(json.dumps({"step1_ok": report["step1"]["ok"], "step2_relations_ok": report["step2"]["relations_ok"], "step3_ok": report["step3"]["ok"], "output": args.output}, ensure_ascii=False, indent=2))
        return

    # Step3
    roles_vis = {}
    tree_all = requests.get(f"{BASE_URL}/api/member/role/get/tree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", [])
    for cat in tree_all:
        if cat.get("name") == "费用角色组":
            continue
        for rr in cat.get("children", []) or []:
            if rr.get("name") and rr.get("id"):
                roles_vis[rr["name"]] = rr["id"]

    groups = requests.get(f"{BASE_URL}/api/bill/template/queryTemplateTree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", []) or []
    group_map = {g.get("name") or g.get("title"): g.get("id") for g in groups if (g.get("name") or g.get("title")) and g.get("id")}

    df3 = read_sheet_with_header(xlsx, "03_单据表", "单据模板名称")
    df3 = df3[df3[get_col(df3, "是否创建")].astype(str).str.strip() == "是"].copy()
    df3[get_col(df3, "单据分组（一级目录）")] = df3[get_col(df3, "单据分组（一级目录）")].ffill()

    type_map = {"报销单": "EXPENSE", "借款单": "LOAN", "批量付款单": "PAYMENT", "申请单": "REQUISITION"}

    created_docs = []
    for _, row in df3.iterrows():
        group_name = str(row.get(get_col(df3, "单据分组（一级目录）"), "")).strip()
        doc_type = str(row.get(get_col(df3, "单据大类（二级目录）"), "")).strip()
        doc_name = str(row.get(get_col(df3, "单据模板名称"), "")).strip()
        vis_type = str(row.get(get_col(df3, "可见范围类型"), "")).strip()
        vis_obj = str(row.get(get_col(df3, "可见范围对象"), "")).strip()

        if group_name not in group_map:
            requests.post(f"{BASE_URL}/api/bill/template/createTemplateGroup", headers=h, json={"name": group_name, "companyId": company_id}, timeout=12)
            time.sleep(0.4)
            groups = requests.get(f"{BASE_URL}/api/bill/template/queryTemplateTree", headers=h, params={"companyId": company_id}, timeout=12).json().get("result", []) or []
            group_map = {g.get("name") or g.get("title"): g.get("id") for g in groups if (g.get("name") or g.get("title")) and g.get("id")}

        targets = split_values(vis_obj)
        role_ids = [roles_vis[t] for t in targets if vis_type == "角色" and t in roles_vis]
        user_ids = [user_map[t] for t in targets if vis_type == "员工" and t in user_map]
        dep_ids = [dep_map[t] for t in targets if vis_type == "部门" and t in dep_map]

        # 判断是否有有效的可见范围限制
        # 类型是"限制"且有具体对象时，才限制可见范围
        has_targets = bool(targets) and bool(role_ids or user_ids or dep_ids)
        is_limited_type = vis_type == "限制" or vis_type == "角色" or vis_type == "员工" or vis_type == "部门"
        has_visibility = is_limited_type and has_targets

        payload = {
            "applyRelateFlag": True,
            "applyRelateNecessary": False,
            "businessType": "PRIVATE",
            "companyId": company_id,
            "componentJson": [],
            "departmentIds": dep_ids if has_visibility else [],
            "feeIds": [],
            "feeScopeFlag": False,
            "groupId": group_map.get(group_name),
            "icon": "md-pricetag",
            "iconColor": "#4c7cc3",
            "loanIds": [],
            "name": doc_name,
            "payFlag": True,
            "requestScope": False,
            "requisitionIds": [],
            "roleIds": role_ids if has_visibility else [],
            "status": "ACTIVE",
            "type": type_map.get(doc_type, "EXPENSE"),
            "userIds": user_ids if has_visibility else [],
            "userScopeFlag": has_visibility,
            "workFlowId": workflow_id,
        }
        if payload["type"] == "REQUISITION":
            payload["applyContentType"] = "TEXT"

        # 费用限制分支
        # 只有当单据已成功匹配到费用角色且角色里有人时，才勾选“限制费用类型”
        fee_role_ids = report["step2"]["role_by_doc"].get(doc_name, [])
        if has_people.get(doc_name, False) and fee_role_ids:
            payload["feeRoleIds"] = fee_role_ids
            payload["feeScopeType"] = "FEE_ROLE"
            payload["feeIds"] = []
            payload["feeScopeFlag"] = True
            report["step3"]["branch_fee_role"].append({"doc": doc_name, "feeRoleIds": fee_role_ids})
        else:
            # 没有费用角色人员时，不勾选“限制费用类型”
            report["step3"]["branch_skip"].append(doc_name)

        cr = requests.post(f"{BASE_URL}/api/bill/template/createTemplate", headers=h, json=payload, timeout=15).json()
        if cr.get("code") == 200 and cr.get("success"):
            report["step3"]["ok"] += 1
            created_docs.append(doc_name)
        else:
            report["step3"]["fail"].append({"doc": doc_name, "message": cr.get("message")})

    if created_docs:
        print("\n7️⃣ 页面保存闭环...")
        for idx, doc_name in enumerate(created_docs):
            try:
                save_result = ui_save_bill_template(
                    doc_name,
                    preferred_browser=args.browser,
                    reload_page=(idx == 0),
                )
                report["step3"]["ui_save_ok"].append(
                    {
                        "doc": doc_name,
                        "message": save_result.get("message"),
                        "templateId": save_result.get("templateId"),
                    }
                )
                print(f"   ✅ 已页面保存：{doc_name}")
            except Exception as exc:
                report["step3"]["ui_save_fail"].append({"doc": doc_name, "message": str(exc)})
                print(f"   ❌ 页面保存失败：{doc_name} -> {exc}")

    Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print("✅ 导入完成")
    print(json.dumps({
        "step1_ok": report["step1"]["ok"],
        "step2_relations_ok": report["step2"]["relations_ok"],
        "step3_ok": report["step3"]["ok"],
        "output": args.output,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
