# -*- coding: utf-8 -*-
"""双击运行即可同步涌益现货数据到云端"""
import shutil, subprocess, os, sys
from pathlib import Path
from datetime import datetime

DESKTOP = Path(r"D:\CC\Desktop")
PROJECT = Path(r"D:\CC\test-claude\sentiment_platform")
TARGET = PROJECT / "data" / "涌益咨询日度数据.xlsx"

print("=" * 50)
print("  涌益现货数据 -> 云端同步")
print("=" * 50)
print()

# 1. 找桌面最新文件
candidates = []
for f in DESKTOP.glob("*涌益咨询*.xlsx"):
    if not f.name.startswith("~$"):
        candidates.append((f.stat().st_mtime, f))

if not candidates:
    print("[错误] 桌面没有涌益咨询 Excel 文件！")
    input("按回车退出...")
    sys.exit(1)

latest = sorted(candidates, key=lambda x: x[0], reverse=True)[0][1]
print(f"[1/3] 找到: {latest.name}")

# 2. 复制到项目
TARGET.parent.mkdir(parents=True, exist_ok=True)
try:
    TARGET.unlink()
except FileNotFoundError:
    pass
shutil.copy2(latest, TARGET)
print(f"[2/3] 已复制到项目")

# 3. 提交并推送
os.chdir(str(PROJECT))
subprocess.run(["git", "add", "data/涌益咨询日度数据.xlsx"], capture_output=True)
r = subprocess.run(
    ["git", "commit", "-m", f"更新现货数据 {datetime.now().strftime('%Y-%m-%d')}"],
    capture_output=True, text=True,
)
output = (r.stdout or "") + (r.stderr or "")
if "nothing to commit" in output:
    print("[3/3] 数据无变化，跳过推送")
else:
    subprocess.run(["git", "push"])
    print("[3/3] 已推送到云端")

print()
print("完成! Streamlit Cloud 将自动更新。")
try:
    input("按回车退出...")
except EOFError:
    pass
