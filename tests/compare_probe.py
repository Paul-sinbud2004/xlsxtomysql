#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""探针样本双版本对照：同一份特殊表格，Python 版与 Rust 版各跑一遍，
逐字节比对 CREATE TABLE 与 INSERT 部分，并比对退出码。

用例表直接复用 xlsxtomysql/tests/probe_run.py 里的 CASES，避免两处维护。

    PY=/path/to/python python3 tests/compare_probe.py [样本名过滤子串 ...]

仅用于开发/自测。
"""
import importlib.util
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                       # 仓库根
WS = os.path.dirname(ROOT)                         # 工作目录（可能没有 Python 实现）
SAMPLES = os.environ.get("SAMPLES") or os.path.join(WS, "xlsxtomysql", "tests", "probe")
PY_VERSION = os.environ.get("PY_VERSION") or os.path.join(WS, "xlsxtomysql", "xlsxtomysql.py")
CASES_PY = os.path.join(WS, "xlsxtomysql", "tests", "probe_run.py")
RS = os.path.join(ROOT, "xlsxtomysql")
OUT = os.path.join(HERE, "cmpprobe")

PY = os.environ.get("PY") or sys.executable

# 开发期检查：需要一个并行的 Python 实现与它的探针用例表。
# 公开发布版仓库不含这些，找不到时明确提示并跳过。
if not os.path.exists(PY_VERSION) or not os.path.exists(CASES_PY):
    print("== 跳过：未找到并行的 Python 实现与探针用例表 ==")
    print(f"   需要 {PY_VERSION}")
    print(f"   需要 {CASES_PY}")
    print("   这是开发期检查，公开发布版仓库不含 Python 实现。")
    print("   想跑的话：PY_VERSION=... CASES_PY=... SAMPLES=... python3 tests/compare_probe.py")
    sys.exit(0)


def load_cases():
    path = CASES_PY
    spec = importlib.util.spec_from_file_location("probe_run", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.CASES


def extract(path):
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        txt = f.read()
    out = []
    keep = False
    for line in txt.splitlines():
        if line.startswith("CREATE TABLE"):
            keep = True
        if keep:
            out.append(line)
            if line.startswith(") ENGINE"):
                keep = False
        elif line.startswith("INSERT"):
            out.append(line)
    return "\n".join(out) + "\n" if out else ""


def main():
    pats = sys.argv[1:]
    os.makedirs(OUT, exist_ok=True)
    cases = load_cases()
    if pats:
        # 两个字母的写法按"组前缀"精确匹配（dt 只匹配 dt01…），避免
        # "dt" 顺带命中 fullwi*d*t*h 之类的样本名；更长的写法按子串匹配。
        def hit(name: str) -> bool:
            group = name.split("_", 1)[0]
            for p in pats:
                if len(p) == 2:
                    if group.startswith(p):
                        return True
                elif p in name:
                    return True
            return False
        cases = [c for c in cases if hit(c[0])]

    ok = bad = 0
    fails = []
    for (fname, sheet, table, hrow, drow, extra) in cases:
        path = os.path.join(SAMPLES, fname)
        if not os.path.exists(path):
            print(f"  [SKIP] {fname}（样本不存在）")
            continue
        base = [path, sheet, table, str(hrow), str(drow), "--progress", "off", "--no-color"]
        # 每个样本各用一份输出文件，不删旧文件（沙箱对删除次数有限制）
        py_sql = os.path.join(OUT, "py_%s.sql" % fname.rsplit(".", 1)[0])
        rs_sql = os.path.join(OUT, "rs_%s.sql" % fname.rsplit(".", 1)[0])
        # 先把旧产物清空（截断而不是删除，避免触发沙箱的删除配额）
        for f in (py_sql, rs_sql):
            open(f, "w").close()
        p1 = subprocess.run([PY, PY_VERSION] + base + ["--out", py_sql] + extra,
                            capture_output=True, text=True, timeout=300)
        p2 = subprocess.run([RS] + base + ["--out", rs_sql] + extra,
                            capture_output=True, text=True, timeout=300)
        a, b = extract(py_sql), extract(rs_sql)
        if p1.returncode != p2.returncode:
            bad += 1
            fails.append(f"{fname}: 退出码不同 python={p1.returncode} rust={p2.returncode}")
            print(f"  [FAIL] {fname}  退出码 python={p1.returncode} rust={p2.returncode}")
            # 打印出错方的第一行非空输出，便于区分"程序问题"与"沙箱限制"
            for who, p in (("py", p1), ("rs", p2)):
                if p.returncode != 0:
                    for ln in ((p.stdout or "") + (p.stderr or "")).splitlines():
                        if ln.strip():
                            print(f"         {who}: {ln.strip()[:140]}")
                            break
        elif a != b:
            bad += 1
            fails.append(f"{fname}: SQL 不一致")
            print(f"  [FAIL] {fname}  SQL 不一致")
            al, bl = a.splitlines(), b.splitlines()
            for i in range(max(len(al), len(bl))):
                x = al[i] if i < len(al) else "<缺行>"
                y = bl[i] if i < len(bl) else "<缺行>"
                if x != y:
                    print(f"         py: {x[:150]}")
                    print(f"         rs: {y[:150]}")
                    break
        else:
            ok += 1
            print(f"  [PASS] {fname}  exit={p1.returncode}")

    print(f"\n================ 通过 {ok} / 失败 {bad} ================")
    if fails:
        print("失败明细：")
        for f in fails:
            print("  - " + f)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
