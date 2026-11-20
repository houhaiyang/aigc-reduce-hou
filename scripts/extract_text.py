#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多格式文档 → 纯文本提取器（aigc-reduce-hou）

把 docx / pdf / txt 三种输入统一转成 UTF-8 纯文本，供 aigc_scan.py 扫描。

依赖（运行时已预置，见 SKILL.md）：
  - docx → python-docx（读段落 + 表格单元格）
  - pdf  → 系统 pdftotext（poppler，优先在 PATH 里找，其次按常见安装路径探测）
  - txt  → 直读（UTF-8，失败退回 gb18030）

用法:
    python extract_text.py <input.{docx|pdf|txt}> [-o output.txt]

不带 -o 时打印到 stdout；带 -o 时写入文件（文件不存在则创建，存在则覆盖）。
"""

import sys
import subprocess
import tempfile
import shutil
from pathlib import Path


# ─── pdftotext 候选路径 ───
PDFTOTEXT_CANDIDATES = [
    shutil.which("pdftotext"),
    r"E:\RJBase\texlive\2026\bin\windows\pdftotext.exe",
    r"C:\Program Files\poppler\Library\bin\pdftotext.exe",
    r"C:\Program Files (x86)\poppler\Library\bin\pdftotext.exe",
]


def find_pdftotext():
    """返回可用的 pdftotext 绝对路径；找不到返回 None。"""
    for cand in PDFTOTEXT_CANDIDATES:
        if cand and Path(cand).exists():
            return str(cand)
    return None


def extract_docx(path: Path) -> str:
    """docx → 纯文本（段落 + 表格，按出现顺序拼接）。"""
    try:
        import docx  # python-docx
    except ImportError as exc:  # pragma: no cover
        print(f"错误: 缺少 python-docx，无法解析 docx — {exc}", file=sys.stderr)
        sys.exit(1)

    document = docx.Document(str(path))
    parts = []
    # 按文档 body 顺序遍历块级元素，保留段落与表格的相对顺序
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    body = document.element.body
    for child in body.iterchildren():
        if child.tag.endswith('}p'):
            parts.append(Paragraph(child, document).text)
        elif child.tag.endswith('}tbl'):
            tbl = Table(child, document)
            for row in tbl.rows:
                cells = [c.text.strip() for c in row.cells]
                parts.append('\t'.join(cells))
    text = '\n'.join(p for p in parts if p.strip())
    return text


def extract_pdf(path: Path) -> str:
    """pdf → 纯文本（经 pdftotext，UTF-8）。"""
    exe = find_pdftotext()
    if not exe:
        print(
            "错误: 未找到 pdftotext（poppler）。请安装 poppler 并把 pdftotext 加入 PATH，"
            "或更新 extract_text.py 里的 PDFTOTEXT_CANDIDATES。",
            file=sys.stderr,
        )
        sys.exit(1)

    # 输出到临时文件再读回，避免 pdftotext 的 stdout 编码问题
    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tf:
        tmp_path = tf.name
    try:
        proc = subprocess.run(
            [exe, "-enc", "UTF-8", "-layout", str(path), tmp_path],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            print(
                f"错误: pdftotext 失败（exit {proc.returncode}）— {proc.stderr.strip()}",
                file=sys.stderr,
            )
            sys.exit(1)
        return Path(tmp_path).read_text(encoding="utf-8")
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def extract_txt(path: Path) -> str:
    """txt → 纯文本（优先 UTF-8，失败退回 gb18030）。"""
    raw = path.read_bytes()
    for enc in ("utf-8", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def extract(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return extract_docx(path)
    if suffix == ".pdf":
        return extract_pdf(path)
    if suffix == ".txt":
        return extract_txt(path)
    # 未知后缀：当纯文本兜底
    print(f"提示: 未知后缀 {suffix!r}，按纯文本读取。", file=sys.stderr)
    return extract_txt(path)


def main():
    if len(sys.argv) < 2:
        print(__doc__, file=sys.stderr)
        sys.exit(2)

    src = Path(sys.argv[1])
    if not src.exists():
        print(f"错误: 文件不存在 — {src}", file=sys.stderr)
        sys.exit(1)

    out_path = None
    if "-o" in sys.argv:
        i = sys.argv.index("-o")
        if i + 1 >= len(sys.argv):
            print("错误: -o 后面缺输出文件路径", file=sys.stderr)
            sys.exit(2)
        out_path = Path(sys.argv[i + 1])

    text = extract(src)

    if out_path:
        out_path.write_text(text, encoding="utf-8")
        print(f"已写入 {out_path}（{len(text)} 字符）", file=sys.stderr)
    else:
        sys.stdout.write(text)


if __name__ == "__main__":
    main()
