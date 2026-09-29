#!/usr/bin/env python3
"""检查「代码支持的选项」与「--help / --man 里写到的选项」是否同步。

发布前最容易出的错不是代码坏了，而是加了选项忘了写文档：
用户看 --help 找不到，或者照文档传了个已经改名的参数。
这里把 parse_args 里认识的选项名抓出来，逐个确认四份文档
（HELP_ZH / HELP_EN / MAN_ZH / MAN_EN）都提到了。

用法：
    python3 tools/check_help_sync.py xlsxtomysql.rs
退出码 0 = 全同步，1 = 有缺漏（缺漏清单打到 stdout）。
"""
import re
import sys

PATH = sys.argv[1] if len(sys.argv) > 1 else "xlsxtomysql.rs"
SRC = open(PATH, encoding="utf-8").read()

# 不需要写进帮助的选项：工具自身用的调试开关，或别名
INTERNAL = {
    "--man-pager",              # 保留
}
ALIASES = {
    "--progress": {"--progress", "--no-progress"},
    "--help": {"--help", "-h"},
    "--version": {"--version", "-V"},
}


def raw_block(name):
    """取出 const NAME: &str = r#"..."#; 之间的内容。"""
    m = re.search(r'const %s: &str = r#"(.*?)"#;' % name, SRC, re.S)
    if not m:
        raise SystemExit("找不到常量 " + name)
    return m.group(1)


def supported_options():
    """parse_args 里 match name.as_str() 的分支 + 直接比较的 -h/-V 等长选项。"""
    opts = set()
    # match name.as_str() { "out" => ..., "lang" => ... }
    m = re.search(r"match name\.as_str\(\) \{(.*?)\n            \}", SRC, re.S)
    if not m:
        raise SystemExit("找不到选项 match 块")
    for lit in re.findall(r'"([a-zA-Z0-9][a-zA-Z0-9-]*)"\s*=>', m.group(1)):
        opts.add("--" + lit)
    # if a == "--xxx" / a == "-h"
    for lit in re.findall(r'a == "(-{1,2}[a-zA-Z0-9][a-zA-Z0-9-]*)"', SRC):
        opts.add(lit)
    # needs_value 之类的表里出现的名字
    for lit in re.findall(r'"(out|err-file|report|max-ident|print-errors)"', SRC):
        opts.add("--" + lit)
    return opts


def main():
    opts = supported_options()
    docs = {n: raw_block(n) for n in ("HELP_ZH", "HELP_EN", "MAN_ZH", "MAN_EN")}
    missing = {}
    for name in sorted(opts):
        if name in INTERNAL:
            continue
        forms = ALIASES.get(name, {name})
        for doc, text in docs.items():
            if not any(f in text for f in forms):
                missing.setdefault(name, []).append(doc)
    print("代码支持的选项: %d 个" % len(opts))
    for name in sorted(opts):
        if name in INTERNAL:
            continue
        forms = sorted(ALIASES.get(name, {name}))
        bad = missing.get(name, [])
        flag = "OK " if not bad else "缺 "
        print("  %s %-18s %s" % (flag, "/".join(forms), "" if not bad else "未出现在: " + ", ".join(bad)))
    if missing:
        print("\n✗ %d 个选项没写进文档" % len(missing))
        return 1
    print("\n✓ help / man 与代码同步")
    return 0


if __name__ == "__main__":
    sys.exit(main())
