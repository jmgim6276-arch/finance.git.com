#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
XLSX=""
OUTPUT=""
BROWSER="edge"
USERNAME_ARG=""
PASSWORD_ARG=""
COMPANY_ID_ARG=""
COMPANY_NAME_ARG=""
TASK_ID=""
AUTO_UPDATE=0
INSTALL_DEPS=0
KEEP_BROWSER=0

usage() {
  cat <<'EOF'
用法：
  bash run_openclaw_import.sh --xlsx "/path/to/客户模板.xlsx" [选项]

选项：
  --xlsx PATH          必填，三表 Excel 路径
  --output PATH        可选，导入报告输出路径
  --browser NAME       可选，edge|chrome|auto，默认 edge
  --username VALUE     可选，财税通手机号；不传则优先用环境变量 CST_USERNAME
  --password VALUE     可选，财税通密码；不传则优先用环境变量 CST_PASSWORD
  --company-id VALUE   可选，多企业账号时指定 companyId；不传则优先用 CST_COMPANY_ID
  --company-name VALUE 可选，期望进入的集团/公司名称；用于校验和多企业切换
  --task-id VALUE      可选，任务隔离 ID；不传则自动生成独立浏览器实例标识
  --keep-browser       可选，导入完成后保留浏览器，不自动关闭
  --update             可选，执行前显式 git 拉取最新代码
  --install            可选，执行前显式安装/校验依赖
  --no-update          可选，兼容旧参数；当前默认就不拉取
  --skip-install       可选，兼容旧参数；当前默认就不安装
  --help               显示帮助

环境变量：
  CST_USERNAME
  CST_PASSWORD
  CST_COMPANY_ID

本地配置文件：
  可在仓库根目录创建 .openclaw.env，脚本会自动加载
EOF
}

ensure_local_no_proxy() {
  local suffix="localhost,127.0.0.1"
  if [[ -n "${NO_PROXY:-}" ]]; then
    export NO_PROXY="${NO_PROXY},${suffix}"
  else
    export NO_PROXY="${suffix}"
  fi
  if [[ -n "${no_proxy:-}" ]]; then
    export no_proxy="${no_proxy},${suffix}"
  else
    export no_proxy="${suffix}"
  fi
}

disable_proxy_for_python_step() {
  unset HTTPS_PROXY HTTP_PROXY ALL_PROXY https_proxy http_proxy all_proxy
}

default_task_id() {
  printf 'task-%s-%s' "$(date +%Y%m%d_%H%M%S)" "$$"
}

sanitize_task_id() {
  printf '%s' "$1" | tr -cs 'A-Za-z0-9._-' '-' | sed 's/^-*//; s/-*$//' | cut -c1-80
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --xlsx)
      XLSX="${2:-}"
      shift 2
      ;;
    --output)
      OUTPUT="${2:-}"
      shift 2
      ;;
    --browser)
      BROWSER="${2:-}"
      shift 2
      ;;
    --username)
      USERNAME_ARG="${2:-}"
      shift 2
      ;;
    --password)
      PASSWORD_ARG="${2:-}"
      shift 2
      ;;
    --company-id)
      COMPANY_ID_ARG="${2:-}"
      shift 2
      ;;
    --company-name)
      COMPANY_NAME_ARG="${2:-}"
      shift 2
      ;;
    --task-id)
      TASK_ID="${2:-}"
      shift 2
      ;;
    --keep-browser)
      KEEP_BROWSER=1
      shift
      ;;
    --update)
      AUTO_UPDATE=1
      shift
      ;;
    --install)
      INSTALL_DEPS=1
      shift
      ;;
    --no-update)
      AUTO_UPDATE=0
      shift
      ;;
    --skip-install)
      INSTALL_DEPS=0
      shift
      ;;
    --help|-h)
      usage
      exit 0
      ;;
    *)
      echo "未知参数: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

if [[ -f "$REPO_DIR/.openclaw.env" ]]; then
  set -a
  # shellcheck disable=SC1091
  source "$REPO_DIR/.openclaw.env"
  set +a
fi

if [[ -z "$XLSX" ]]; then
  echo "缺少 --xlsx 参数" >&2
  usage >&2
  exit 2
fi

if [[ ! -f "$XLSX" ]]; then
  echo "Excel 文件不存在: $XLSX" >&2
  exit 2
fi

if [[ -z "$TASK_ID" ]]; then
  TASK_ID="$(default_task_id)"
fi
TASK_ID="$(sanitize_task_id "$TASK_ID")"
if [[ -z "$TASK_ID" ]]; then
  TASK_ID="$(default_task_id)"
fi

if [[ -z "$OUTPUT" ]]; then
  OUTPUT="$REPO_DIR/agent2_import_report_${TASK_ID}_$(date +%Y%m%d_%H%M%S).json"
fi

cd "$REPO_DIR"
ensure_local_no_proxy()

if [[ "$AUTO_UPDATE" -eq 1 ]]; then
  if git rev-parse --is-inside-work-tree >/dev/null 2>&1 && git remote get-url origin >/dev/null 2>&1; then
    echo "==> 拉取最新 main..."
    git fetch origin main
    git pull --ff-only origin main
  else
    echo "==> 当前目录未配置 origin，跳过自动更新"
  fi
fi

echo "==> 当前版本: $(git log -1 --oneline 2>/dev/null || echo '非 git 仓库')"
echo "==> 任务隔离 ID: $TASK_ID"

if [[ "$INSTALL_DEPS" -eq 1 ]]; then
  echo "==> 安装/校验依赖..."
  python3 -m pip install -r "$REPO_DIR/requirements.txt"
fi

disable_proxy_for_python_step

CMD=(
  python3
  "$REPO_DIR/scripts/import_from_agent1.py"
  --xlsx "$XLSX"
  --output "$OUTPUT"
  --auto-login
  --browser "$BROWSER"
  --task-id "$TASK_ID"
)

if [[ -n "$USERNAME_ARG" ]]; then
  CMD+=(--username "$USERNAME_ARG")
fi

if [[ -n "$PASSWORD_ARG" ]]; then
  CMD+=(--password "$PASSWORD_ARG")
fi

if [[ -n "$COMPANY_ID_ARG" ]]; then
  CMD+=(--company-id "$COMPANY_ID_ARG")
fi

if [[ -n "$COMPANY_NAME_ARG" ]]; then
  CMD+=(--company-name "$COMPANY_NAME_ARG")
fi

echo "==> 开始导入..."
"${CMD[@]}"
echo "==> 导入完成，报告文件: $OUTPUT"

if [[ "$KEEP_BROWSER" -eq 1 ]]; then
  echo "==> 已按要求保留浏览器，不执行自动关闭"
  exit 0
fi

if python3 - "$OUTPUT" <<'PY'
import json
import sys
from pathlib import Path

report_path = Path(sys.argv[1])
if not report_path.exists():
    print("导入报告不存在，跳过自动关闭")
    raise SystemExit(1)

report = json.loads(report_path.read_text(encoding="utf-8"))
checks = [
    ("step1.fail", report.get("step1", {}).get("fail", [])),
    ("step1_department_sync.fail", report.get("step1_department_sync", {}).get("fail", [])),
    ("step1_roles.fail", report.get("step1_roles", {}).get("fail", [])),
    ("step2.relations_fail", report.get("step2", {}).get("relations_fail", [])),
    ("step2.reset_fail", report.get("step2", {}).get("reset_fail", [])),
    ("step3.fail", report.get("step3", {}).get("fail", [])),
    ("step3.default_model_fail", report.get("step3", {}).get("default_model_fail", [])),
    ("step3.ui_save_fail", report.get("step3", {}).get("ui_save_fail", [])),
]
failed = [(name, value) for name, value in checks if value]
if failed:
    print("导入报告仍包含失败项：")
    for name, value in failed:
        print(f"- {name}: {len(value)}")
    raise SystemExit(1)
print("导入报告校验通过，可自动关闭浏览器")
PY
then
  echo "==> 导入结果校验通过，关闭财税通自动化浏览器..."
  if bash "$REPO_DIR/run_openclaw_close_browser.sh" --browser "$BROWSER" --task-id "$TASK_ID" --no-update; then
    echo "==> 导入成功，浏览器已关闭"
  else
    echo "==> 导入成功，但浏览器关闭失败，请检查上面的错误信息" >&2
    exit 1
  fi
else
  echo "==> 导入报告中仍有失败项，保留浏览器供排查"
fi
