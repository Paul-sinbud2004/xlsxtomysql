#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 tools/i18n_catalog.py 的译文表写进 xlsxtomysql.rs，并做覆盖率复核。

用法:
    python3 tools/apply_i18n.py --check     # 只检查（占位符数量、重复键、覆盖率）
    python3 tools/apply_i18n.py --write     # 生成 CATALOG 常量并替换进 xlsxtomysql.rs
"""
import importlib.util
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# 这些问题不在 CATALOG 里的字符串，是按语言分支硬编码拼出来的（带 :.0 之类的格式说明），
# 或者是分隔符常量本身，不需要走表。列在这里是为了让覆盖率检查有据可依。
EXEMPT = {
    "{:.0} 毫秒", "{:.1} 秒", "{:.0} 分 {:.0} 秒", "{:.0} 时 {:.0} 分",  # fmt_dur
    "、", "；",  # sep() / sep2()
}

# 内嵌 Python 助手里这些字符串不是界面文案，不需要译文（扫描时按源码字面量比对）：
EXEMPT_HELPER = {
    # strftime 解析模板
    "%Y年%m月%d日 %H:%M:%S", "%Y年%m月%d日 %H:%M", "%Y年%m月%d日", "%Y年%m月%d",
    # 中文列名/值的提示词（匹配数据用，翻译反而会误伤数据）
    "日期", "出生", "生日", "时间", "年月日", "[年月日]",
    "、", "；",
}

BEGIN = "// ==== I18N CATALOG BEGIN (由 tools/apply_i18n.py 生成，勿手改) ===="
END = "// ==== I18N CATALOG END ===="


def load_pairs():
    spec = importlib.util.spec_from_file_location(
        "i18n_catalog", os.path.join(HERE, "i18n_catalog.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return list(mod.PAIRS)


def placeholders(s):
    # 只有 {} 算占位符；{:.0} / {:5.1} 这类带格式说明的不算（它们不参与模板匹配）
    return s.count("{}")


def validate(pairs):
    errs = []
    seen = {}
    for zh, en in pairs:
        if zh in seen:
            errs.append("重复的中文键: %r" % zh)
        seen[zh] = en
        if placeholders(zh) != placeholders(en):
            errs.append("占位符数量不一致 (%d vs %d): %r -> %r"
                        % (placeholders(zh), placeholders(en), zh, en))
    return errs, seen


def rust_lit(s):
    out = ['"']
    for ch in s:
        if ch == "\\":
            out.append("\\\\")
        elif ch == '"':
            out.append('\\"')
        elif ch == "\n":
            out.append("\\n")
        elif ch == "\r":
            out.append("\\r")
        elif ch == "\t":
            out.append("\\t")
        elif ch == "\x1b":
            out.append("\\x1b")
        elif ord(ch) < 0x20:
            out.append("\\u{%x}" % ord(ch))
        else:
            out.append(ch)
    out.append('"')
    return "".join(out)


def render_catalog(pairs):
    # 先按"是否含占位符"分组，再按中文键排序，输出稳定、便于 diff
    exact = sorted([(z, e) for z, e in pairs if "{}" not in z], key=lambda x: x[0])
    tmpl = sorted([(z, e) for z, e in pairs if "{}" in z],
                  key=lambda x: (-sum(len(s) for s in x[0].split("{}")), x[0]))
    lines = [BEGIN,
             "/// 译文表：中文原文 → 英文。键里 {} 是运行期填入的值。",
             "/// 生成自 tools/i18n_catalog.py（改文案请改那里再跑 tools/apply_i18n.py --write）。",
             "const CATALOG: &[(&str, &str)] = &[",
             "    // ---- 整串精确匹配 ----"]
    for z, e in exact:
        lines.append("    (%s, %s)," % (rust_lit(z), rust_lit(e)))
    lines.append("    // ---- 模板匹配（整串锚定；长字面量优先，避免短模板抢先命中）----")
    for z, e in tmpl:
        lines.append("    (%s, %s)," % (rust_lit(z), rust_lit(e)))
    lines.append("];")
    lines.append(END)
    return "\n".join(lines)


def strip_python(src):
    """剥掉 Python 注释与三引号文档串（助手里的三引号全是 docstring）。"""
    out = list(src)
    i, n = 0, len(src)
    while i < n:
        if src.startswith('"""', i) or src.startswith("'''", i):
            q = src[i:i + 3]
            e = src.find(q, i + 3)
            e = n if e == -1 else e + 3
            for k in range(i, e):
                if out[k] != "\n":
                    out[k] = " "
            i = e
            continue
        if src[i] == "#":
            e = src.find("\n", i)
            e = n if e == -1 else e
            for k in range(i, e):
                out[k] = " "
            i = e
            continue
        if src[i] in '"\'':
            q = src[i]
            k = i + 1
            while k < n and src[k] != q:
                if src[k] == "\\":
                    k += 1
                k += 1
            i = k + 1
            continue
        i += 1
    return "".join(out)


PY_FMT = re.compile(r"%(?:\d+\$)?[-+ #0]*\d*(?:\.\d+)?[sdrifgxXeE%]")


def helper_messages(path):
    """内嵌助手里"会被用户看到"的中文文案：已剥注释/文档串，% 格式已换成 {}。"""
    import re as _re
    src = open(path, encoding="utf-8").read()
    i = src.index('const PY_HELPER: &str = r##"')
    j = src.index('"##;', i)
    body = src[i + len('const PY_HELPER: &str = r##"'):j]
    clean = strip_python(body)
    cjk = _re.compile(r"[\u2000-\u206f\u2190-\u21ff\u3000-\u303f"
                      r"\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]")
    out = {}
    for ln, line in enumerate(clean.split("\n"), 1):
        for m in _re.finditer(r'"((?:[^"\\]|\\.)*)"|\'((?:[^\'\\]|\\.)*)\'', line):
            s = m.group(1) if m.group(1) is not None else m.group(2)
            if not s or not cjk.search(s):
                continue
            norm = PY_FMT.sub("{}", s)
            if norm in EXEMPT_HELPER or s in EXEMPT_HELPER:
                continue
            out.setdefault(norm, ln)
    return out


def inventory(path):
    spec = importlib.util.spec_from_file_location(
        "i18n_inventory", os.path.join(HERE, "i18n_inventory.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    rust = mod.collect(path, helper_only=False)
    return set(rust), set(helper_messages(path))


def main():
    pairs = load_pairs()
    errs, seen = validate(pairs)
    for e in errs:
        print("✗ " + e)
    if errs:
        return 1

    target = os.path.join(ROOT, "xlsxtomysql.rs")
    rust_lits, helper_lits = inventory(target)

    missing_rust = sorted(x for x in rust_lits if x not in seen and x not in EXEMPT)
    missing_helper = sorted(x for x in helper_lits if x not in seen and x not in EXEMPT)
    unused = sorted(k for k in seen if k not in rust_lits and k not in helper_lits)

    print("译文表条目: %d（精确 %d / 模板 %d）"
          % (len(pairs),
             sum(1 for z, _ in pairs if "{}" not in z),
             sum(1 for z, _ in pairs if "{}" in z)))
    print("源码里的中文文案: Rust %d 条 / Python 助手 %d 条" % (len(rust_lits), len(helper_lits)))
    if missing_rust:
        print("\n✗ Rust 侧缺译文 %d 条:" % len(missing_rust))
        for m in missing_rust:
            print("   %r" % m)
    if missing_helper:
        print("\n✗ 助手侧缺译文 %d 条:" % len(missing_helper))
        for m in missing_helper:
            print("   %r" % m)
    if unused:
        print("\n· 译文表里暂时没被源码引用的键 %d 条（可能已改成 t() 调用，属正常）:" % len(unused))
        for u in unused[:20]:
            print("   %r" % u)
        if len(unused) > 20:
            print("   ... 其余 %d 条" % (len(unused) - 20))

    if "--write" in sys.argv:
        src = open(target, encoding="utf-8").read()
        block = render_catalog(pairs)
        a = src.index(BEGIN)
        b = src.index(END) + len(END)
        src = src[:a] + block + src[b:]
        open(target, "w", encoding="utf-8").write(src)
        print("\n已写入 CATALOG：%d 条" % len(pairs))

    if missing_rust or missing_helper:
        return 1
    print("\n✓ 覆盖率检查通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
