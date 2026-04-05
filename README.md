# agent2 - 财税通三表导入 Skill（小白可用）

这个仓库是给小白 AI/新同学直接上手的：

- 输入：Agent1 生成的三表 Excel（01/02/03）
- 输出：按 Step1 + Step2 + Step2.5 + Step3 导入财税通系统

## 1) 先决条件

1. 准备好 Agent1 生成的 `.xlsx`
2. 本机已安装 Edge 或 Chrome
3. 如果想让脚本自动登录，准备好财税通手机号和密码

## 2) 一键执行

```bash
python3 -m pip install -r requirements.txt
python3 scripts/preflight_check.py
python3 scripts/import_from_agent1.py --xlsx "/path/to/三表联动_客户模板_xxxx.xlsx" --output "./agent2_import_report.json"
```

如果你希望导入时自动打开浏览器并登录，可以直接这样跑：

```bash
python3 scripts/import_from_agent1.py \
  --xlsx "/path/to/三表联动_客户模板_xxxx.xlsx" \
  --output "./agent2_import_report.json" \
  --auto-login \
  --username "你的手机号"
```

脚本会在终端里隐藏输入密码，不建议把密码写进命令行历史。

## 2.5) 新设备快速开始

```bash
git clone https://github.com/jmgim6276-arch/finance.git.com.git
cd finance.git.com
python3 -m pip install -r requirements.txt
python3 scripts/ensure_browser_login.py --username "你的手机号"
python3 scripts/preflight_check.py --auto-login --username "你的手机号"
```

- 如果你在 macOS 上使用系统自带 Python，看见 `urllib3` 的 LibreSSL 提示通常不影响脚本执行
- 如果你想让环境更稳，优先使用 Homebrew 或 python.org 安装的 Python 3.9+

## 3) 成功标准（小白版）

- preflight 全绿
- step1/step2/step3 有成功计数
- 报告文件生成成功
- 失败项有明确原因

## 4) 这份副本已包含的修复

- 兼容 Sheet2/Sheet3 里带数字污染的表头，例如 `归属单01据名称`
- 创建费用科目时补齐详情字段，避免只生成树节点、不生成可读模板详情
- 费用角色绑定优先按完整姓名匹配，并支持自动创建费用类型角色
- 费用科目名称匹配会自动 `strip()`，降低尾部空格导致的一级科目识别失败
- 支持在导入前自动打开 Edge/Chrome，并通过终端隐藏输入密码自动登录财税通
- 支持登录后自动选择企业；如果是多企业账号，可传 `--company-id`
- Step3 创建单据模板后，会自动打开单据模板页并执行一次页面保存，补齐浏览器侧闭环

## 4.5) 自动登录建议

推荐方式：

```bash
export CST_USERNAME="你的手机号"
python3 scripts/ensure_browser_login.py
```

如果是多企业账号：

```bash
export CST_USERNAME="你的手机号"
export CST_COMPANY_ID="8108"
python3 scripts/ensure_browser_login.py
```

密码默认会在终端隐藏输入。

也可以用环境变量直接传：

```bash
export CST_USERNAME="你的手机号"
export CST_PASSWORD="你的密码"
python3 scripts/preflight_check.py --auto-login
```

不推荐把密码直接写在命令行里，因为 shell 历史可能会保留。

## 5) 先看这两份文档

- `references/README.md`

## 6) 关键边界（必须遵守）

- 只按最新流程：Step1 + Step2 + Step2.5 + Step3
- Step3 不创建审批流，只引用 workflowId
- 借款单/申请单不进入费用限制分支
- Sheet2 `归属单据名称` 与 Sheet3 `单据模板名称` 必须一一对应

---

如果你是小白，按文档和命令一步步走，不要跳步骤、不要猜字段。
