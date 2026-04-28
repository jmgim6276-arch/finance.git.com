#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BROWSER="auto"
AUTO_UPDATE=0
DRY_RUN=0
TIMEOUT="5"

usage() {
  cat <<'EOF'
用法：
  bash run_openclaw_close_browser.sh [选项]

选项：
  --browser NAME       可选，auto|edge|chrome，默认 auto
  --timeout SECONDS    可选，关闭后验证秒数，默认 5
  --dry-run            可选，只检测，不实际关闭
  --update             可选，执行前显式 git 拉取最新代码
  --no-update          可选，兼容旧参数；当前默认就不拉取
  --help               显示帮助
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --browser)
      BROWSER="${2:-}"
      shift 2
      ;;
    --timeout)
      TIMEOUT="${2:-}"
      shift 2
      ;;
    --dry-run)
      DRY_RUN=1
      shift
      ;;
    --update)
      AUTO_UPDATE=1
      shift
      ;;
    --no-update)
      AUTO_UPDATE=0
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

cd "$REPO_DIR"

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

CMD=(
  python3
  "$REPO_DIR/scripts/close_cst_browser.py"
  --browser "$BROWSER"
  --timeout "$TIMEOUT"
)

if [[ "$DRY_RUN" -eq 1 ]]; then
  CMD+=(--dry-run)
fi

echo "==> 关闭财税通自动化浏览器..."
"${CMD[@]}"
echo "==> 关闭流程完成"
