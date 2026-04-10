#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
XLSX=""
OUTPUT=""
BROWSER="edge"
USERNAME_ARG=""
PASSWORD_ARG=""
COMPANY_ID_ARG=""
AUTO_UPDATE=1
INSTALL_DEPS=1

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
  --no-update          可选，跳过 git 拉取最新代码
  --skip-install       可选，跳过 pip 安装依赖
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

if [[ -z "$OUTPUT" ]]; then
  OUTPUT="$REPO_DIR/agent2_import_report_$(date +%Y%m%d_%H%M%S).json"
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

echo "==> 开始导入..."
"${CMD[@]}"
echo "==> 导入完成，报告文件: $OUTPUT"
