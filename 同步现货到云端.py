#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
涌益现货数据 → 云端一键同步
用法：双击运行，或在终端执行 python 同步现货到云端.py
"""
import os
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

DESKTOP = Path(r"D:\CC\Desktop")
PROJECT = Path(r"D:\CC\test-claude\sentiment_platform")
DATA_DIR = PROJECT / "data"
TARGET = DATA_DIR / "涌益咨询日度数据.xlsx"


def find_latest_spot():
    """在桌面找最新的涌益咨询 Excel"""
    candidates = []
    for pattern in ["*涌益咨询日度数据*.xlsx", "*涌益咨询*.xlsx"]:
        for f in DESKTOP.glob(pattern):
            # 跳过 Office 临时文件
            if f.name.startswith("~$"):
                continue
            if f not in [c[1] for c in candidates]:
                candidates.append((f.stat().st_mtime, f))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def main():
    print("=" * 60)
    print("  涌益现货数据 → 云端一键同步")
    print("=" * 60)
    print()

    # 1. 找最新文件
    latest = find_latest_spot()
    if latest is None:
        print("[错误] 桌面上没有找到涌益咨询 Excel 文件！")
        print("请把文件放到桌面，文件名包含「涌益咨询」即可。")
        input("按回车退出...")
        return 1

    print(f"[1/3] 找到最新文件: {latest.name}")
    print(f"       路径: {latest}")

    # 2. 复制到项目（先删后写，避免文件锁定）
    print(f"[2/3] 复制到项目目录...")
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    try:
        TARGET.unlink()
    except FileNotFoundError:
        pass
    shutil.copy2(latest, TARGET)
    print(f"       → {TARGET}")

    # 3. 提交并推送
    print(f"[3/3] 提交并推送到 GitHub...")
    os.chdir(PROJECT)

    subprocess.run(["git", "add", str(TARGET.relative_to(PROJECT))], check=True)
    today = datetime.now().strftime("%Y年%m月%d日")
    result = subprocess.run(
        ["git", "commit", "-m", f"更新涌益现货数据至{today}"],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    output = (result.stdout or "") + (result.stderr or "")
    if "nothing to commit" in output:
        print("       (数据未变化，跳过提交)")
    else:
        print(f"       提交成功")

    subprocess.run(["git", "push"], check=True)

    print()
    print("=" * 60)
    print("  同步完成！Streamlit Cloud 将自动重新部署。")
    print("=" * 60)
    input("按回车退出...")
    return 0


if __name__ == "__main__":
    exit(main())
