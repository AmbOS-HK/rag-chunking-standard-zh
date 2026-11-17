#!/usr/bin/env python3
"""CI 用同步校验：docs/ 下 docx 必须与 scripts/ 生成脚本保持一致（单一事实源强制检查）。

流程：快照 docs/*.docx 正文文本 → npm run build:all 重新生成 → 再次提取正文 → 逐文件比对。
任何文件正文漂移、构建失败、文件集合变化均判失败。

用法: python tools/check_sync.py   （在仓库根目录执行，需先 npm ci）
退出码: 0 = 全部同步; 1 = 存在漂移或构建失败。
"""
import glob
import os
import re
import subprocess
import sys
import zipfile


def text_of(path: str):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)


def snapshot():
    return {
        p: text_of(p)
        for p in glob.glob("docs/*.docx")
        if not os.path.basename(p).startswith("~$")
    }


def main() -> int:
    before = snapshot()
    if not before:
        print("FAIL docs/ 下无可校验的 docx")
        return 1
    r = subprocess.run(
        "npm run build:all", shell=True, capture_output=True,
        encoding="utf-8", errors="replace",
    )
    if r.returncode != 0:
        print("FAIL npm run build:all 构建失败：")
        print((r.stdout or "")[-2000:])
        print((r.stderr or "")[-2000:])
        return 1
    after = snapshot()
    ok = True
    for p in sorted(set(before) | set(after)):
        if p not in before:
            print(f"FAIL {p}: 构建产生了快照中不存在的新文件（文件集合变化）")
            ok = False
        elif p not in after:
            print(f"FAIL {p}: 构建后文件消失")
            ok = False
        elif before[p] != after[p]:
            print(f"FAIL {p}: 正文与生成脚本不一致——请改 scripts/ 后重新生成，勿直接编辑 docx")
            ok = False
        else:
            print(f"SYNC {p}")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
