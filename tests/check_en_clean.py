#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""英文模式残留检查：控制台输出与产物里不允许出现中文。

原理：用 tests/enfix 里生成的全英文样例（表名/工作表名/字段名/数据都是英文），
所以 --lang=en 时输出里出现的**任何**中文字符都只能是漏翻的界面文案。

用法:
    python3 tests/check_en_clean.py [二进制路径]

退出码 0 = 干净；1 = 发现残留（逐条列出文件与行）。
"""
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
BIN = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "xlsxtomysql")
WORK = os.path.join(HERE, "enfix_out")
FIX = os.path.join(HERE, "enfix")

CJK = re.compile(r"[\u2000-\u206f\u2190-\u21ff\u3000-\u303f\u4e00-\u9fff"
                 r"\uf900-\ufaff\uff00-\uffef]")
# 全角/中文标点在英文里可以接受的特殊项（目前没有，留作扩展位）
ALLOW = []

RUNS = [
    # (样例, sheet, 表名, 表头行, 首数据行, 附加参数, 期望退出码集合)
    ("en_clean.xlsx", "Data", "students", 1, 2, [], {0}),
    ("en_clean.xlsx", "Data", "scan_only", 1, 2, ["--scan-only"], {0}),
    ("en_clean.xlsx", "Data", "sampled", 1, 2,
     ["--sample-threshold", "10", "--sample-ratio", "0.2"], {0}),
    ("en_clean.xlsx", "Data", "report", 1, 2, ["--report", "REPORT"], {0}),
    ("en_clean.xlsx", "Data", "ddl", 1, 2,
     ["--drop-table", "--add-id", "id", "--not-null", "--insert-ignore",
      "--table-comment", "students"], {0}),
    ("en_fail.xlsx", "Data", "fails", 1, 2, [], {0}),
    ("en_messy.xlsx", "Data", "messy", 1, 2, [], {2}),
    ("en_messy.xlsx", "Data", "messy_force", 1, 2, ["--force"], {0}),
    ("en_merge.xlsx", "Data", "merged", 1, 3, [], {0}),
    ("en_merge.xlsx", "Data", "merged_nofill", 1, 3, ["--no-fill-down"], {0}),
]


def run():
    if os.path.isdir(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK)
    if not os.path.isdir(FIX):
        subprocess.run([sys.executable, os.path.join(HERE, "make_en_fixture.py"), FIX],
                       check=True)
    bad = []
    for name, sheet, table, hrow, drow, extra, codes in RUNS:
        src = os.path.join(FIX, name)
        out_sql = os.path.join(WORK, table + ".sql")
        want_report = "REPORT" in extra
        extra = [a for a in extra if a not in ("REPORT", "--report")]
        args = [BIN, src, sheet, table, str(hrow), str(drow),
                "--lang=en", "--out", out_sql, "--err-file",
                os.path.join(WORK, table + "_err.xlsx")] + extra
        if want_report:
            args += ["--report", os.path.join(WORK, table + ".txt")]
        p = subprocess.run(args, capture_output=True, text=True, errors="replace")
        if p.returncode not in codes:
            bad.append((table, "退出码", "期望 %s 实得 %d" % (sorted(codes), p.returncode)))
        for tag, text in (("stdout", p.stdout), ("stderr", p.stderr)):
            for ln, line in enumerate(text.split("\n"), 1):
                if CJK.search(line) and not any(a in line for a in ALLOW):
                    bad.append((table, tag, "第 %d 行: %s" % (ln, line.rstrip())))
        # 产物：.sql / 报表
        for f in [out_sql, os.path.join(WORK, table + ".txt")]:
            if not os.path.exists(f):
                continue
            with open(f, encoding="utf-8", errors="replace") as fh:
                for ln, line in enumerate(fh, 1):
                    if CJK.search(line) and not any(a in line for a in ALLOW):
                        bad.append((table, os.path.basename(f), "第 %d 行: %s" % (ln, line.rstrip())))
    return bad


if __name__ == "__main__":
    problems = run()
    if not problems:
        print("✓ 英文模式无中文残留（%d 组用例，含 stdout/stderr/.sql/--report）" % len(RUNS))
        sys.exit(0)
    print("✗ 发现 %d 处中文残留：" % len(problems))
    for t, where, msg in problems:
        print("  [%s/%s] %s" % (t, where, msg))
    sys.exit(1)
