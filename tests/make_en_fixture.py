#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成"全英文表头"的样例，用来验证英文模式下没有任何中文残留。

思路：如果样例的表名、工作表名、字段名、数据全是英文，那么 --lang=en 时
控制台与产物里**任何**中文都只能是漏翻的界面文案 —— 这样就能用
"输出里是否有 CJK 字符"做自动化断言，不需要人工逐条比对。

生成：
    en_clean.xlsx   干净表：各种类型都覆盖（整数/小数/日期/时间/日期时间/布尔/长文本/空列）
    en_messy.xlsx   脏表：分隔行 + 重复表头 + 中间空行 + 合计行 + 说明行
    en_fail.xlsx    会产出失败行的表：varchar 长度不足 + 整数溢出 + 非法日期
    en_merge.xlsx   表头纵向合并 + 数据区纵向合并 + 隐藏列 + 百分比格式 + 时区偏移
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import numbers
from openpyxl.utils import get_column_letter

OUT = sys.argv[1] if len(sys.argv) > 1 else "tests/enfix"
os.makedirs(OUT, exist_ok=True)


def save(wb, name):
    path = os.path.join(OUT, name)
    wb.save(path)
    print("  %-16s -> %s" % (name, path))


def clean():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Full Name", "Age", "Height m", "Birth Date", "Arrive Time",
               "Note", "Lunch Fee", "Boarding", "Recorded At", "Empty Col"])
    for i in range(1, 61):
        ws.append([
            "Student %d" % i,
            8 + (i % 8),
            round(1.20 + (i % 12) * 0.07, 2),
            "2024-03-%02d" % (1 + (i % 28)),
            "0%d:%02d:00" % (7 + (i % 2), i % 60),
            "note-%d" % i,
            round(5 + (i % 8) * 0.25, 2),
            (i % 2 == 0),
            "2024-03-%02d 08:30:00" % (1 + (i % 28)),
            None,
        ])
    save(wb, "en_clean.xlsx")


def messy():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Full Name", "Age", "Score"])
    ws.append(["Ann", 12, 95.5])
    ws.append(["Bob", 13, 88.0])
    ws.append(["----------", "----", "-----"])
    ws.append(["Full Name", "Age", "Score"])
    ws.append(["Cindy", 11, 91.0])
    ws.append([None, None, None])
    ws.append(["Dave", 10, 77.5])
    ws.append(["Total", None, 352.0])
    save(wb, "en_messy.xlsx")


def fails():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Code", "Tiny", "When"])
    for i in range(1, 21):
        ws.append(["A%03d" % i, 100 + i, "2024-04-%02d" % (1 + i % 28)])
    ws.append(["this-code-is-way-too-long-for-varchar-16", 200, "2024-04-30"])
    ws.append(["OK", 99999, "2024-04-30"])
    ws.append(["OK", 5, "not-a-date"])
    save(wb, "en_fail.xlsx")


def merge():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Grade"
    ws.merge_cells("A1:A2")
    ws["B1"] = "Full Name"
    ws["C1"] = "Score"
    ws["D1"] = "Starts"
    ws["E1"] = "Ratio"
    ws["F1"] = "Hidden"
    rows = [["Ann", 91.0, "2024-05-01T08:30:00+08:00", 0.125, "x"],
            [None, 82.5, "2024-05-02T09:00:00+08:00", 0.5, "y"],
            [None, 77.0, "2024-05-03T10:15:00+08:00", 0.875, "z"],
            ["Bob", 68.0, "2024-05-04T11:45:00+08:00", 0.25, "w"]]
    for r in rows:
        ws.append(r)
    ws.merge_cells("A3:A5")
    ws["A3"] = "Group 1"
    ws.column_dimensions["F"].hidden = True
    for cell in ("E3", "E4", "E5", "E6"):
        ws[cell].number_format = "0.0%"
    save(wb, "en_merge.xlsx")


if __name__ == "__main__":
    clean()
    messy()
    fails()
    merge()
