#!/usr/bin/env python3
"""扫描 xlsxtomysql.rs，列出所有"面向用户的"中文字符串（已按 Rust 转义还原）。

用途：i18n 覆盖率复核。英文模式下，每一条都必须能从译文表（catalog）里查到，
否则用户会看到中英混杂。这个工具负责把源码里的中文文案一条不漏地挖出来。

用法:
    python3 tools/i18n_inventory.py xlsxtomysql.rs            # 人类可读清单
    python3 tools/i18n_inventory.py xlsxtomysql.rs --json     # JSON（给生成器/校验器用）
    python3 tools/i18n_inventory.py xlsxtomysql.rs --helper   # 只列内嵌 Python 助手部分

实现要点：
  * 先剥离注释（// 与 /* */），再词法扫描字符串字面量；
  * 正确还原 Rust 转义：\\n \\t \\r \\\\ \\" \\0 \\xNN \\u{...} 以及行尾 \\ 续行（吃掉换行与行首空白）；
  * PY_HELPER / HELP_ZH / MAN_ZH 三个语言资源常量会先被排除；
  * 关键词表（ID_HINTS / MONEY_HINTS / SUBTOTAL_HINTS）命中中文列名，属数据匹配逻辑，
    不是界面文案，排除。
"""
import json
import re
import sys

CJK = re.compile(r"[\u2000-\u206f\u2190-\u21ff\u3000-\u303f\u3400-\u4dbf"
                  r"\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]")

PY_HELPER_OPEN = 'const PY_HELPER: &str = r##"'
PY_HELPER_CLOSE = '"##;'
RESOURCE_CONSTS = [
    (PY_HELPER_OPEN, PY_HELPER_CLOSE),
    ('const HELP_ZH: &str = r#"', '"#;'),
    ('const MAN_ZH: &str = r#"', '"#;'),
]
KEYWORD_ARRAYS = ["ID_HINTS", "MONEY_HINTS", "SUBTOTAL_HINTS"]


def blank_out(src, spans):
    out = list(src)
    for a, b in spans:
        for i in range(a, b):
            if out[i] != "\n":
                out[i] = " "
    return "".join(out)


def const_spans(src):
    spans = []
    for pat, close in RESOURCE_CONSTS:
        start = src.index(pat)
        end = src.index(close, start) + len(close)
        spans.append((start, end))
    return spans


def keyword_spans(src):
    spans = []
    for name in KEYWORD_ARRAYS:
        m = re.search(r"const %s: &\[&str\] = &\[" % re.escape(name), src)
        if not m:
            continue
        end = src.index("];", m.end()) + 2
        spans.append((m.start(), end))
    return spans


def strip_comments(src):
    res = list(src)
    i, n = 0, len(src)
    state = None
    depth = 0
    while i < n:
        c = src[i]
        nxt = src[i + 1] if i + 1 < n else ""
        if state is None:
            if c == "/" and nxt == "/":
                state = "line"
                i += 2
                continue
            if c == "/" and nxt == "*":
                state = "block"
                depth = 1
                i += 2
                continue
        elif state == "line":
            if c == "\n":
                state = None
            else:
                res[i] = " "
            i += 1
            continue
        elif state == "block":
            if c == "/" and nxt == "*":
                depth += 1
                res[i] = res[i + 1] = " "
                i += 2
                continue
            if c == "*" and nxt == "/":
                depth -= 1
                res[i] = res[i + 1] = " "
                i += 2
                if depth == 0:
                    state = None
                continue
            if c != "\n":
                res[i] = " "
            i += 1
            continue
        i += 1
    return "".join(res)


SIMPLE_ESC = {
    "n": "\n", "t": "\t", "r": "\r", "\\": "\\", '"': '"', "'": "'",
    "0": "\0", "a": "\x07", "b": "\x08", "f": "\x0c", "v": "\x0b",
}


def unescape_rust(s):
    out = []
    i, n = 0, len(s)
    while i < n:
        c = s[i]
        if c != "\\":
            out.append(c)
            i += 1
            continue
        nxt = s[i + 1] if i + 1 < n else ""
        if nxt in SIMPLE_ESC:
            out.append(SIMPLE_ESC[nxt])
            i += 2
            continue
        if nxt == "x" and re.fullmatch(r"[0-9a-fA-F]{2}", s[i + 2:i + 4]):
            out.append(chr(int(s[i + 2:i + 4], 16)))
            i += 4
            continue
        if nxt == "u" and i + 2 < n and s[i + 2] == "{":
            end = s.index("}", i + 3)
            out.append(chr(int(s[i + 3:end], 16)))
            i = end + 1
            continue
        if nxt == "\n":
            i += 2
            while i < n and s[i] in " \t":
                i += 1
            continue
        out.append(nxt)
        i += 2
    return "".join(out)


def literals(src):
    found = []
    i, n = 0, len(src)
    line = 1
    while i < n:
        c = src[i]
        if c == "\n":
            line += 1
            i += 1
            continue
        j = i + 1 if (c == "b" and src[i + 1:i + 2] == "r") else i
        if src[j:j + 1] == "r":
            k = j + 1
            hashes = 0
            while k < n and src[k] == "#":
                hashes += 1
                k += 1
            if k < n and src[k] == '"':
                close = '"' + "#" * hashes
                end = src.find(close, k + 1)
                if end == -1:
                    break
                found.append((line, src[k + 1:end]))
                line += src.count("\n", i, end + len(close))
                i = end + len(close)
                continue
        if c == '"':
            j = i + 1
            buf = []
            while j < n:
                if src[j] == "\\":
                    buf.append(src[j:j + 2])
                    j += 2
                    continue
                if src[j] == '"':
                    break
                if src[j] == "\n":
                    line += 1
                buf.append(src[j])
                j += 1
            found.append((line, unescape_rust("".join(buf))))
            i = j + 1
            continue
        if c == "'" and re.match(r"'(\\.|[^\\'])'", src[i:]):
            i += re.match(r"'(\\.|[^\\'])'", src[i:]).end()
            continue
        i += 1
    return found


def collect(path, helper_only=False):
    src = open(path, encoding="utf-8").read()
    if helper_only:
        start = src.index(PY_HELPER_OPEN) + len(PY_HELPER_OPEN)
        end = src.index(PY_HELPER_CLOSE, start)
        body = src[start:end]
    else:
        body = strip_comments(blank_out(src, const_spans(src) + keyword_spans(src)))
    seen = {}
    for ln, s in literals(body):
        if CJK.search(s):
            seen.setdefault(s, ln)
    return seen


def main():
    args = list(sys.argv[1:])
    mode = "text"
    for flag in ("--json", "--helper"):
        if flag in args:
            mode = flag[2:]
            args.remove(flag)
    path = args[0] if args else "xlsxtomysql.rs"
    seen = collect(path, helper_only=(mode == "helper"))
    if mode == "json":
        print(json.dumps(seen, ensure_ascii=False, indent=1, sort_keys=True))
        return
    print("含中文的字符串字面量：%d 条" % len(seen))
    for s in sorted(seen, key=lambda x: seen[x]):
        print("%5d  %s" % (seen[s], s.replace("\n", "\\n").replace("\t", "\\t")))


if __name__ == "__main__":
    main()
