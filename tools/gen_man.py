#!/usr/bin/env python3
"""从 xlsxtomysql.rs 的 MAN_EN / MAN_ZH 常量生成 man page（roff）。

目的：保证「`--man` 打印的手册」与「安装的 man page」永远是同一份内容，
不会出现两处各写一遍、改了一处忘另一处的情况。

用法：
    python3 tools/gen_man.py xlsxtomysql.rs            # 生成 docs/xlsxtomysql.1（英文）
    python3 tools/gen_man.py xlsxtomysql.rs --lang zh  # 生成 docs/xlsxtomysql.zh.1
    python3 tools/gen_man.py xlsxtomysql.rs -o -       # 打到 stdout
"""
import argparse
import os
import re
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# 手册里的一级小节 → roff .SH 名称
SECTION_MAP = {
    "Name": "NAME",
    "Usage": "SYNOPSIS",
    "Positional arguments": "ARGUMENTS",
    "Options": "OPTIONS",
    "Exit codes": "EXIT STATUS",
    "Type inference": "TYPE INFERENCE",
    "Sampled skip scanning": "SAMPLING",
    "Special sheet layouts": "SPECIAL SHEET LAYOUTS",
    "Reports": "REPORTS",
    "Environment variables": "ENVIRONMENT",
    "Examples": "EXAMPLES",
    "Limitations": "LIMITATIONS",
    "License": "LICENSE",
    "名称": "NAME",
    "用法": "SYNOPSIS",
    "位置参数": "ARGUMENTS",
    "选项": "OPTIONS",
    "退出码": "EXIT STATUS",
    "类型推断": "TYPE INFERENCE",
    "抽样扫描": "SAMPLING",
    "特殊表格": "SPECIAL SHEET LAYOUTS",
    "报表": "REPORTS",
    "环境变量": "ENVIRONMENT",
    "示例": "EXAMPLES",
    "已知限制": "LIMITATIONS",
    "许可": "LICENSE",
}


def roff_escape(text):
    """roff 里 \\ 是转义引导符，- 在行首会被当成连字符命令，都要处理。"""
    text = text.replace("\\", "\\e")
    text = text.replace("-", "\\-")
    if text.startswith("'") or text.startswith("."):
        text = "\\&" + text
    return text


def extract(src_path, const):
    src = open(src_path, encoding="utf-8").read()
    m = re.search(r'const %s: &str = r#"(.*?)"#;' % const, src, re.S)
    if not m:
        raise SystemExit("× 源码里找不到常量 " + const)
    return m.group(1)


def version_of(src_path):
    src = open(src_path, encoding="utf-8").read()
    m = re.search(r'const VERSION: &str = "([^"]+)"', src)
    return m.group(1) if m else "unknown"


def to_roff(manual, version, lang):
    """把「纯文本手册」整段塞进 roff 的 .nf/.fi 里（不做重排，保持 --man 的原始版式）。"""
    raw = manual.split("\n")
    # 去掉标题行与 ===== 分隔线
    body_lines = []
    for line in raw:
        if set(line.strip()) == {"="} and line.strip():
            continue
        body_lines.append(line)
    while body_lines and not body_lines[0].strip():
        body_lines.pop(0)
    if body_lines and re.search(r"(manual|手册)\s*$", body_lines[0]):
        body_lines.pop(0)
    while body_lines and not body_lines[0].strip():
        body_lines.pop(0)

    # 从 Name / 名称 段里取 NAME 行，并把这一段从正文里摘掉（避免出现两个 .SH NAME）
    name_desc = ""
    cleaned = []
    i = 0
    while i < len(body_lines):
        line = body_lines[i]
        stripped = line.strip()
        if SECTION_MAP.get(stripped) == "NAME" and line == stripped:
            i += 1
            for l2 in body_lines[i:]:
                if l2.strip():
                    name_desc = l2.strip()
                    i += 1
                    break
                i += 1
            continue
        cleaned.append(line)
        i += 1

    out = []
    out.append('.\\" 由 tools/gen_man.py 从 xlsxtomysql.rs 的 %s 常量生成，勿手改'
               % ("MAN_EN" if lang == "en" else "MAN_ZH"))
    out.append('.TH XLSXTOMYSQL 1 "%s" "xlsxtomysql %s" "User Commands"'
               % (time.strftime("%Y-%m-%d"), version))
    out.append(".SH NAME")
    if not name_desc:
        name_desc = ("xlsxtomysql \\- convert Excel into MySQL statements" if lang == "en"
                     else "xlsxtomysql \\- 把 Excel 转成 MySQL 语句")
    out.append(roff_escape(name_desc).replace("\\e-", "\\-"))

    # 正文：第一个小节标题之前的零散内容归到 DESCRIPTION，避免出现空的 .SH
    pending = []
    in_pre = False
    started = False
    for line in cleaned:
        stripped = line.strip()
        sec = SECTION_MAP.get(stripped)
        if sec and line == stripped and sec != "NAME":
            if not started:
                if any(x.strip() for x in pending):
                    out.append(".SH DESCRIPTION")
                    out.append(".nf")
                    out.extend(roff_escape(x) for x in pending)
                    out.append(".fi")
                started = True
            if in_pre:
                out.append(".fi")
            out.append(".SH " + sec)
            out.append(".nf")
            in_pre = True
            continue
        if not started:
            pending.append(line)
            continue
        if not in_pre:
            out.append(".nf")
            in_pre = True
        out.append(roff_escape(line))
    if in_pre:
        out.append(".fi")
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?", default=os.path.join(ROOT, "xlsxtomysql.rs"))
    ap.add_argument("--lang", choices=["en", "zh"], default="en")
    ap.add_argument("-o", "--out", default=None)
    args = ap.parse_args()

    const = "MAN_EN" if args.lang == "en" else "MAN_ZH"
    manual = extract(args.src, const)
    version = version_of(args.src)
    text = to_roff(manual, version, args.lang)

    if args.out == "-":
        sys.stdout.write(text)
        return
    dest = args.out or os.path.join(ROOT, "docs",
                                    "xlsxtomysql.1" if args.lang == "en" else "xlsxtomysql.zh.1")
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    open(dest, "w", encoding="utf-8").write(text)
    print("✓ 已生成 %s（%d 行，来自 %s）" % (dest, text.count("\n"), const))


if __name__ == "__main__":
    main()
