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

如果你想给 OpenClaw 或别的电脑一个固定入口，直接用仓库根目录这个脚本：

```bash
bash run_openclaw_import.sh --xlsx "/path/to/三表联动_客户模板_xxxx.xlsx"
```

这条脚本会自动做几件事：

- 默认直接使用当前本地仓库版本
- 默认不主动 `git pull`、不主动安装依赖
- 导入时强制带 `--auto-login`
- 浏览器登录态过期时，自动重新登录财税通
- 导入报告没有失败项时，自动关闭财税通自动化浏览器

如果你明确要在执行前更新代码或补装依赖，再额外加：

```bash
bash run_openclaw_import.sh --xlsx "/path/to/三表联动_客户模板_xxxx.xlsx" --update --install
```

如果你不想每次手动输入账号密码，可以在仓库根目录放一个 `.openclaw.env`，格式参考 `.openclaw.env.example`。

如果你这次想保留浏览器方便排查：

```bash
bash run_openclaw_import.sh --xlsx "/path/to/三表联动_客户模板_xxxx.xlsx" --keep-browser
```

如果你更希望每次都临时发给 OpenClaw 不同账号，也可以直接这样跑：

```bash
bash run_openclaw_import.sh \
  --xlsx "/path/to/三表联动_客户模板_xxxx.xlsx" \
  --username "你的手机号" \
  --password "你的密码" \
  --company-id "8108" \
  --company-name "上海公司"
```

这种方式不会要求你固定使用 `.openclaw.env`。更适合你有多个财税通账号轮流登录的场景。
如果你只有集团/公司名称，没有 `company-id`，也建议至少传 `--company-name`，脚本会用它校验并切换目标企业，避免误复用上一次别的公司浏览器上下文。

## 2.2) 只登录财税通，不做导入

如果你只想让 OpenClaw 先登录财税通，不和任何导入项目绑定，直接用这个独立入口：

```bash
bash run_openclaw_login.sh \
  --username "本次要登录的手机号" \
  --password "本次要登录的密码" \
  --company-id "8108"
```

这个脚本只做三件事：

- 默认直接使用当前本地仓库版本
- 默认不主动 `git pull`、不主动安装依赖
- 自动打开浏览器并登录财税通

如果你明确要在登录前更新代码或补装依赖，再加 `--update --install`。

它不会读取 Excel，也不会执行 Step1/Step2/Step3。

## 2.3) 只关闭财税通自动化浏览器

如果你只想把 `finance.git.com` 这套自动化开的 Edge/Chrome 关掉，用这个入口：

```bash
bash run_openclaw_close_browser.sh --browser auto
```

这个脚本只会处理财税通自动化浏览器实例：

- 默认识别 `~/.finance-cst/edge-cdp-profile`
- 默认识别 `~/.finance-cst/chrome-cdp-profile`
- 只在进程消失且本地 CDP 端口不可用后，才返回关闭成功

默认也不会主动更新仓库；如果你确实需要，再显式加 `--update`。

如果你只是想先确认它会关到哪个浏览器，不真正执行：

```bash
bash run_openclaw_close_browser.sh --dry-run
```

## 2.5) 新设备快速开始

```bash
git clone https://github.com/jmgim6276-arch/finance.git.com.git
cd finance.git.com
cp .openclaw.env.example .openclaw.env
# 然后把 .openclaw.env 里的手机号/密码/companyId 改成你自己的
bash run_openclaw_import.sh --xlsx "/path/to/三表联动_客户模板_xxxx.xlsx"
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
- 支持登录后自动选择企业；如果是多企业账号，建议同时传 `--company-id` 与 `--company-name`
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
- 只会识别这三张 sheet：`01_添加员工`、`02_费用科目配置`、`03_单据表`
- 你在 Excel 后面新增无关 sheet 没关系，脚本会忽略它们

---

如果你是小白，按文档和命令一步步走，不要跳步骤、不要猜字段。
