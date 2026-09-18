# -*- coding: utf-8 -*-
"""
把 Desktop 生成的看板 HTML 同步到两个位置：
  1) 同目录 dist 文件夹  —— 本地预览用（双击 index.html 即可看）
  2) 项目目录 dist 文件夹 —— 云端部署源（保持原有分享链接不变）

由同目录的 run_update.bat 在 build_data.py 之后自动调用，无需手动执行。
"""
import os
import shutil
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent

# 云端部署源（保持不变，才能复用原 sandboxId / 原分享链接）
DIST_DEPLOY = Path(r"C:\Users\Faye\WorkBuddy\2026-08-20-15-44-18\退费数据看板\dist")

TARGETS = [
    ("本地预览", HERE / "dist"),
    ("云端部署源", DIST_DEPLOY),
]

HTMLS = ["dashboard.html", "compare.html"]
ENTRY = "index.html"


def main():
    ok = True
    for label, dist in TARGETS:
        try:
            dist.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"  [{label}] 无法创建目录 {dist}: {e}")
            ok = False
            continue

        for name in HTMLS:
            src = HERE / name
            if not src.exists():
                print(f"  [{label}] 缺少 {name}，跳过")
                ok = False
                continue
            dst = dist / name
            shutil.copy2(src, dst)
            print(f"  [{label}] {name}  {src.stat().st_size:,} B")

        # 入口页：部署源必须有；缺失时从本地补一份
        entry_dst = dist / ENTRY
        if not entry_dst.exists():
            entry_src = HERE / "dist" / ENTRY
            if entry_src.exists():
                shutil.copy2(entry_src, entry_dst)
                print(f"  [{label}] {ENTRY} 已补齐")

    print("同步完成" if ok else "同步完成（有警告，见上）")


if __name__ == "__main__":
    main()
