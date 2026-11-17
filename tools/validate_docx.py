#!/usr/bin/env python3
"""DOCX 结构完整性校验：ZIP 完整性、必备部件存在性、XML 良构性。

用法: python tools/validate_docx.py <file.docx> [...]
退出码: 0 = 全部通过; 1 = 任一文件失败。
"""
import sys
import zipfile
import xml.etree.ElementTree as ET

REQUIRED_PARTS = ["[Content_Types].xml", "_rels/.rels", "word/document.xml"]


def validate(path: str) -> bool:
    ok = True
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        print(f"FAIL {path}: 不是有效 ZIP/docx ({e})")
        return False
    bad = z.testzip()
    if bad:
        print(f"FAIL {path}: ZIP 损坏条目 {bad}")
        ok = False
    names = set(z.namelist())
    for part in REQUIRED_PARTS:
        if part not in names:
            print(f"FAIL {path}: 缺少必备部件 {part}")
            ok = False
    for name in names:
        if name.endswith((".xml", ".rels")):
            try:
                ET.fromstring(z.read(name))
            except ET.ParseError as e:
                print(f"FAIL {path}: XML 非良构 {name} ({e})")
                ok = False
    if ok:
        print(f"PASS {path}（{len(names)} 个部件）")
    return ok


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    sys.exit(0 if all(validate(p) for p in sys.argv[1:]) else 1)
