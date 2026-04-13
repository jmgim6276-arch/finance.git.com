#!/usr/bin/env python3
import argparse
import os
import signal
import socket
import subprocess
import sys
import time
from pathlib import Path

BROWSERS = [
    {
        "id": "edge",
        "name": "Edge",
        "port": 9223,
        "profile_dir": str(Path.home() / ".finance-cst" / "edge-cdp-profile"),
    },
    {
        "id": "chrome",
        "name": "Chrome",
        "port": 18800,
        "profile_dir": str(Path.home() / ".finance-cst" / "chrome-cdp-profile"),
    },
]


def browser_choices(preferred: str):
    if preferred == "edge":
        return [browser for browser in BROWSERS if browser["id"] == "edge"]
    if preferred == "chrome":
        return [browser for browser in BROWSERS if browser["id"] == "chrome"]
    return BROWSERS


def is_port_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex(("127.0.0.1", port)) == 0


def list_browser_processes(browser: dict) -> list[dict]:
    patterns = [
        browser["profile_dir"],
        f"--remote-debugging-port={browser['port']}",
    ]
    matches: dict[int, str] = {}
    result = subprocess.run(
        ["ps", "-axww", "-o", "pid=,command="],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"ps failed for {browser['name']}: {result.stderr.strip()}")
    for raw_line in result.stdout.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        pid_text, _, cmd = line.partition(" ")
        if not pid_text.isdigit():
            continue
        pid = int(pid_text)
        if pid == os.getpid():
            continue
        if any(pattern in cmd for pattern in patterns):
            matches[pid] = cmd
    return [{"pid": pid, "cmd": cmd} for pid, cmd in sorted(matches.items())]


def send_signal(browser: dict, sig: int) -> None:
    for process in list_browser_processes(browser):
        try:
            os.kill(process["pid"], sig)
        except ProcessLookupError:
            continue


def wait_for_browser_exit(browser: dict, timeout: float) -> tuple[bool, list[dict], bool]:
    deadline = time.time() + timeout
    latest_processes = list_browser_processes(browser)
    latest_port_open = is_port_open(browser["port"])
    clear_checks = 0
    while time.time() < deadline:
        latest_processes = list_browser_processes(browser)
        latest_port_open = is_port_open(browser["port"])
        if not latest_processes and not latest_port_open:
            clear_checks += 1
            if clear_checks >= 2:
                return True, latest_processes, latest_port_open
        else:
            clear_checks = 0
        time.sleep(0.25)
    latest_processes = list_browser_processes(browser)
    latest_port_open = is_port_open(browser["port"])
    return (not latest_processes and not latest_port_open), latest_processes, latest_port_open


def format_processes(processes: list[dict]) -> str:
    return "; ".join(f"{item['pid']} {item['cmd']}" for item in processes)


def close_browser(browser: dict, timeout: float, dry_run: bool) -> tuple[bool, str]:
    before_processes = list_browser_processes(browser)
    before_port_open = is_port_open(browser["port"])

    if not before_processes and not before_port_open:
        return True, f"ℹ️ {browser['name']} 财税通自动化浏览器本来就是关闭状态"

    if dry_run:
        parts = [f"🔎 {browser['name']} 检测到待关闭目标"]
        if before_processes:
            parts.append(f"进程: {format_processes(before_processes)}")
        parts.append(f"CDP端口 {browser['port']}: {'open' if before_port_open else 'closed'}")
        return True, " | ".join(parts)

    send_signal(browser, signal.SIGTERM)
    closed, remaining_processes, remaining_port_open = wait_for_browser_exit(
        browser,
        timeout=min(timeout, 3.0),
    )
    if not closed:
        send_signal(browser, signal.SIGKILL)
        closed, remaining_processes, remaining_port_open = wait_for_browser_exit(browser, timeout)

    if closed:
        return True, f"✅ 已关闭 {browser['name']} 财税通自动化浏览器"

    details = []
    if remaining_processes:
        details.append(f"残留进程: {format_processes(remaining_processes)}")
    if remaining_port_open:
        details.append(f"CDP端口 {browser['port']} 仍然可访问")
    detail_text = " | ".join(details) if details else "未通过关闭校验"
    return False, f"❌ 关闭 {browser['name']} 失败: {detail_text}"


def main() -> int:
    parser = argparse.ArgumentParser(description="关闭财税通自动化浏览器并验证结果")
    parser.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome"],
        default="auto",
        help="关闭哪个浏览器；默认 auto",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=5.0,
        help="强制关闭后的验证等待秒数，默认 5",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只检测目标，不实际关闭",
    )
    args = parser.parse_args()

    messages: list[str] = []
    all_ok = True
    for browser in browser_choices(args.browser):
        ok, message = close_browser(browser, timeout=max(args.timeout, 0.5), dry_run=args.dry_run)
        messages.append(message)
        all_ok = all_ok and ok

    for message in messages:
        print(message)

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
