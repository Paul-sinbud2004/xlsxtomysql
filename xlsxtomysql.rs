//! xlsxtomysql —— 把 Excel 转成 MySQL 建表 + 插入语句（Rust 单文件实现）
//!
//! 编译（不需要任何外部 crate）：
//!     rustc -O main.rs -o xlsxtomysql
//!
//! 用法：
//!     xlsxtomysql 文件名.xlsx sheet名 新表名 字段名称所在行 第一个数据所在行 [共几行]
//!
//! 设计要点
//! --------
//! * Rust 负责：参数解析、表格格式检查、类型推断、SQL 生成、报表、进度条。
//! * Excel 的读取（.xlsx / .xls）与 errrows.xlsx 的写出交给内嵌的 Python 助手，
//!   那段源码就放在本文件的 `PY_HELPER` 常量里，随二进制一起编译进去，运行时用
//!   `python3 -c <那段源码>` 调用 —— 磁盘上不会留下任何 .py 文件（`--dump-python`
//!   可把它原样导出来单独调试）。
//! * 两边用制表符分隔的行协议通信：Python 的 stdout 走数据通道、stderr 走进度与
//!   日志通道，Rust 用两个线程分别接管，互不干扰。

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{self, BufRead, BufReader, BufWriter, IsTerminal, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{channel, Receiver};
use std::thread;
use std::time::{Instant, SystemTime, UNIX_EPOCH};

// ===========================================================================
// 内嵌的 Python 助手
// ===========================================================================
const PY_HELPER: &str = r##"#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlsxtomysql 的 Excel 读写助手。

这个文件的全部内容会被逐字内嵌进 xlsxtomysql-rs 的 Rust 源码常量 PY_HELPER，
由宿主以 `python3 -c <本源码> <mode> [参数...]` 的方式调用 —— 运行时不落地成文件。

协议
----
stdout：制表符分隔的文本行。以 # 开头的是控制行，其余为数据行。
    #KIND<TAB>格式名                    xlsx / Excel 2003 (.xls)
    #SHEET<TAB>表名<TAB>总行<TAB>总列
    #SHEETS<TAB>表1<TAB>表2...
    #MERGES<TAB>r1,c1,r2,c2|r1,c1,r2,c2|...
    #FORMULA<TAB>0|1
    #HEADER<TAB>单元格...
    #ROW<TAB>行号<TAB>单元格...
    #END<TAB>读取行数<TAB>最后读到的行号<TAB>已输出行数
    #OK<TAB>绝对路径
    #FATAL<TAB>退出码<TAB>消息        退出码与宿主进程的退出码约定一致
每个单元格编码为 "<kind>:<已转义的值>"，kind 为单字符：
    e 空 / i 整数 / d 小数 / f 科学计数法 / b 布尔 / s 文本
    D 日期 / T 时间 / M 日期时间 / x Excel 错误值
日期时间的值用规范文本：#D:2024-03-05 / T:09:30:00 / M:2024-03-05 09:30:00

stderr：#P<TAB>当前<TAB>总数  是进度，由宿主渲染成进度条；其余行原样转发给用户。
"""

import datetime as dt
import math
import os
import re
import sys
import zipfile

# --------------------------------------------------------------------------
# 协议输出
# --------------------------------------------------------------------------
TAB = "\t"


def w(line):
    sys.stdout.write(line)
    sys.stdout.write("\n")


def progress(n, total):
    sys.stderr.write("#P\t%d\t%d\n" % (n, total))
    sys.stderr.flush()


def note(msg):
    sys.stderr.write(str(msg) + "\n")
    sys.stderr.flush()


_ESC = {"\\": "\\\\", "\t": "\\t", "\n": "\\n", "\r": "\\r"}
_UNESC = {"t": "\t", "n": "\n", "r": "\r", "\\": "\\"}


def esc(s):
    out = []
    for ch in s:
        out.append(_ESC.get(ch, ch))
    return "".join(out)


def unesc(s):
    if "\\" not in s:
        return s
    out = []
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if c == "\\" and i + 1 < n:
            out.append(_UNESC.get(s[i + 1], s[i + 1]))
            i += 2
        else:
            out.append(c)
            i += 1
    return "".join(out)


# --------------------------------------------------------------------------
# 值分类（与 Python 版 xlsxtomysql.py 保持同一套判定规则）
# --------------------------------------------------------------------------
K_EMPTY, K_INT, K_DEC, K_FLOAT = "e", "i", "d", "f"
K_BOOL, K_TEXT = "b", "s"
K_DATE, K_TIME, K_DATETIME = "D", "T", "M"
K_ERROR = "x"
K_OOR = "o"          # 日期/时间超出 MySQL 可表示范围，按文本处理

RE_INT = re.compile(r"^[+-]?\d+$")
RE_DEC = re.compile(r"^[+-]?(?:\d+\.\d*|\.\d+)$")
RE_SCI = re.compile(r"^[+-]?\d+(?:\.\d+)?[eE][+-]?\d+$")
RE_LEADING_ZERO = re.compile(r"^[+-]?0\d")
RE_CONTROL = re.compile(r"[\x00-\x08\x0b-\x0c\x0e-\x1f\x7f]")

# 与 xlsxtomysql.py 保持完全一致：先试更具体的（带毫秒/时区/秒），再试宽松的
DATE_FORMATS_DATETIME = (
    "%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M:%S.%f",
    "%Y-%m-%d %H:%M:%S%z", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
    "%Y.%m.%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M",
    "%Y年%m月%d日 %H:%M:%S", "%Y年%m月%d日 %H:%M",
    "%Y-%m-%d %I:%M:%S %p", "%Y-%m-%d %I:%M %p",
    "%Y/%m/%d %I:%M:%S %p", "%Y/%m/%d %I:%M %p",
)
DATE_FORMATS_DATE = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日", "%Y年%m月%d")
DATE_FORMATS_TIME = ("%H:%M:%S.%f", "%H:%M:%S", "%H:%M",
                     "%I:%M:%S %p", "%I:%M %p", "%I:%M:%S%p", "%I:%M%p")
RE_DURATION = re.compile(r"^([+-]?)(\d{1,3}):([0-5]\d):([0-5]\d)$")
RE_COMPACT_DATE = re.compile(r"^(19|20)\d{6}$")
# 带时区偏移的日期时间：MySQL 的 datetime 不存时区，偏移会被丢弃
RE_TZ_OFFSET = re.compile(r"(?:Z|[+-]\d{2}:?\d{2})$")

MYSQL_DATE_MIN = dt.date(1000, 1, 1)
MYSQL_DATE_MAX = dt.date(9999, 12, 31)
MYSQL_TIME_MAX_SEC = 838 * 3600 + 59 * 60 + 59

# 字段名语义提示：命中则对单元格额外尝试紧凑日期（20240305）
DATE_NAME_HINTS = ("日期", "出生", "生日", "时间", "date", "birth", "day")


def clean(s):
    return RE_CONTROL.sub("", str(s)).strip()


def looks_int(s):
    return RE_INT.match(s) is not None


def looks_dec(s):
    return RE_DEC.match(s) is not None or looks_int(s)


def temporal_in_range(kind, value):
    """该日期/时间能否被 MySQL 的 date/datetime/time 表示。"""
    if kind in (K_DATE, K_DATETIME):
        d = value.date() if isinstance(value, dt.datetime) else value
        if not isinstance(d, dt.date):
            return True
        return MYSQL_DATE_MIN <= d <= MYSQL_DATE_MAX
    if kind == K_TIME:
        if isinstance(value, dt.timedelta):
            secs = value.total_seconds()
        elif isinstance(value, dt.time):
            secs = value.hour * 3600 + value.minute * 60 + value.second
        else:
            return True
        return -MYSQL_TIME_MAX_SEC <= secs <= MYSQL_TIME_MAX_SEC
    return True


def parse_date_like(s, allow_compact=False):
    if len(s) < 4 or len(s) > 32:
        return None
    # 含中文时必须是"年月日"这类日期写法，避免把普通中文当日期
    if re.search(r"[\u4e00-\u9fff]", s) and not re.search(r"[年月日]", s):
        return None
    for f in DATE_FORMATS_DATETIME:
        try:
            d = dt.datetime.strptime(s, f)
            if d.tzinfo is not None:
                d = d.replace(tzinfo=None)     # 时区信息 MySQL 存不下，按字面时间保存
            return K_DATETIME, d
        except ValueError:
            pass
    for f in DATE_FORMATS_DATE:
        try:
            return K_DATE, dt.datetime.strptime(s.strip(), f).date()
        except ValueError:
            pass
    for f in DATE_FORMATS_TIME:
        try:
            return K_TIME, dt.datetime.strptime(s, f).time()
        except ValueError:
            pass
    # 超过 24 小时只能写成"时长"（25:30:00 / 100:00:00），MySQL time 支持
    m = RE_DURATION.match(s)
    if m and int(m.group(2)) >= 24:
        sign = -1 if m.group(1) == "-" else 1
        td = dt.timedelta(hours=int(m.group(2)), minutes=int(m.group(3)),
                          seconds=int(m.group(4)))
        return K_TIME, td * sign
    if allow_compact and RE_COMPACT_DATE.match(s):
        try:
            return K_DATE, dt.datetime.strptime(s, "%Y%m%d").date()
        except ValueError:
            pass
    return None


def cell(kind, value=""):
    return kind, value


CELL_EMPTY = (K_EMPTY, "")


def classify_string(raw):
    s = clean(raw)
    if s == "":
        return CELL_EMPTY
    if RE_SCI.match(s):
        try:
            return (K_FLOAT, repr(float(s)))
        except ValueError:
            pass
    if looks_int(s):
        if RE_LEADING_ZERO.match(s):
            return (K_TEXT, s)                       # 前导零 → 必须当文本
        if len(s.lstrip("+-")) > 18:
            return (K_TEXT, s)                       # 超长数字 → 文本
        try:
            return (K_INT, str(int(s)))
        except ValueError:
            return (K_TEXT, s)
    if looks_dec(s):
        if RE_LEADING_ZERO.match(s.lstrip("+-")):
            return (K_TEXT, s)
        try:
            return (K_DEC, repr(float(s)))
        except ValueError:
            return (K_TEXT, s)
    d = parse_date_like(s)
    if d:
        return cell_from_date_kind(d[0], d[1])
    return (K_TEXT, s)


def cell_from_date_kind(kind, value):
    """日期值 → 行协议。超出 MySQL 范围的一律按文本（K_OOR）输出。"""
    if not temporal_in_range(kind, value):
        if kind == K_DATETIME:
            return (K_OOR, value.strftime("%Y-%m-%d %H:%M:%S"))
        if kind == K_DATE:
            return (K_OOR, value.strftime("%Y-%m-%d"))
        return (K_OOR, fmt_time(value))
    if kind == K_DATETIME:
        return (K_DATETIME, value.strftime("%Y-%m-%d %H:%M:%S"))
    if kind == K_DATE:
        return (K_DATE, value.strftime("%Y-%m-%d"))
    if kind == K_TIME:
        return (K_TIME, fmt_time(value))
    return (K_TEXT, clean(value))


def fmt_time(t):
    if isinstance(t, dt.timedelta):
        secs = int(t.total_seconds())
    elif isinstance(t, dt.time):
        secs = t.hour * 3600 + t.minute * 60 + t.second
    else:
        secs = int(t)
    # 时长要保留完整小时数（MySQL time 支持到 838:59:59），不能对 24 取模
    sign = "-" if secs < 0 else ""
    secs = abs(secs)
    h, rem = divmod(secs, 3600)
    m, s = divmod(rem, 60)
    return "%s%02d:%02d:%02d" % (sign, h, m, s)


def classify_native_date(value, number_format=None):
    fmt = (number_format or "").lower()
    has_hour = ("h" in fmt) or ("[h]" in fmt)
    has_date = bool(re.search(r"[yd]", fmt))
    if isinstance(value, dt.datetime):
        if value.year <= 1900 and ((value.month == 1 and value.day <= 1) or value.year == 1899):
            return CELL_EMPTY
        if has_hour and not has_date:
            return cell_from_date_kind(K_TIME, value.time())
        if has_date and not has_hour:
            return cell_from_date_kind(K_DATE, value.date())
        if not has_date and not has_hour:
            return cell_from_date_kind(K_DATETIME, value)
        if value.hour == value.minute == value.second == 0 and not has_hour:
            return cell_from_date_kind(K_DATE, value.date())
        return cell_from_date_kind(K_DATETIME, value)
    if isinstance(value, dt.date):
        if value.year <= 1900:
            return CELL_EMPTY
        return cell_from_date_kind(K_DATE, value)
    if isinstance(value, dt.time):
        return cell_from_date_kind(K_TIME, value)
    if isinstance(value, dt.timedelta):
        return cell_from_date_kind(K_TIME, value)
    return (K_TEXT, clean(value))


def encode(c):
    kind, val = c
    return kind + ":" + esc(val)


# --------------------------------------------------------------------------
# 格式探测
# --------------------------------------------------------------------------
def sniff_format(path):
    with open(path, "rb") as f:
        head = f.read(8)
    if head[:2] == b"PK":
        return "xlsx"
    if head[:4] == b"\xd0\xcf\x11\xe0":
        return "xls"
    if head[:5] == b"<?xml" or head[:1] == b"<":
        return "xml"
    return "unknown"


def parse_range(ref):
    m = re.fullmatch(r"([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?", ref.strip())
    if not m:
        return None

    def colnum(s):
        n = 0
        for ch in s.upper():
            n = n * 26 + (ord(ch) - 64)
        return n

    c1 = colnum(m.group(1))
    r1 = int(m.group(2))
    c2 = colnum(m.group(3)) if m.group(3) else c1
    r2 = int(m.group(4)) if m.group(4) else r1
    return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


ROW_RE = re.compile(rb"<row[^>]*\br=\"(\d+)\"")
COL_RE = re.compile(rb"<c[^>]*\br=\"([A-Z]{1,3})\d+\"")
DIM_RE = re.compile(rb"<dimension[^>]*\bref=\"([^\"]+)\"")
MERGE_BLOCK_RE = re.compile(rb"<mergeCells[^>]*>(.*?)</mergeCells>", re.S)
MERGE_ONE_RE = re.compile(rb"<mergeCell[^>]*ref=\"([^\"]+)\"")
HCOL_RE = re.compile(rb"<col\b[^>]*hidden=\"(?:1|true)\"[^>]*/?>")
HROW_RE = re.compile(rb"<row\b[^>]*hidden=\"(?:1|true)\"[^>]*>")
ATTR_NUM_RE = re.compile(rb"(\w+)=\"(\d+)\"")

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def xlsx_scan(path, sheet, merge_limit_mb=64, merge_scan="auto"):
    """流式扫描工作表 XML：合并区域 / 公式 / 真实行列数。"""
    import xml.etree.ElementTree as ET

    merges = []
    has_formula = False
    has_embed_img = False
    hidden_cols = []
    hidden_rows = []
    max_row = 0
    max_col = 0
    dim_rows = 0
    dim_cols = 0

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        if "xl/workbook.bin" in names:
            raise HelperError("检测到 .xlsb 二进制工作簿，openpyxl 不支持，请另存为 .xlsx 后重试")
        try:
            wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
            rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        except KeyError:
            raise HelperError("不是有效的 .xlsx 文件（缺少工作簿定义）")
        rels = {r.get("Id"): r.get("Target") for r in rels_xml}
        names_in_wb = []
        target = None
        for sh in wb_xml.iter(NS_MAIN + "sheet"):
            names_in_wb.append(sh.get("name"))
            if sh.get("name") == sheet:
                target = rels.get(sh.get(NS_REL + "id"))
        if target is None:
            raise HelperError(
                "工作表 %r 不存在。可用工作表: %s" % (sheet, "、".join(names_in_wb)),
                EXIT_USAGE)
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        target = target.replace("xl/xl/", "xl/")

        size = os.path.getsize(path)
        do_merge = merge_scan != "off" and (merge_scan == "on" or size <= merge_limit_mb * 1024 * 1024)

        tail = b""
        part = 0
        with z.open(target) as fh:
            while True:
                chunk = fh.read(1 << 20)
                if not chunk:
                    break
                part += 1
                buf = tail + chunk
                if not has_formula and (b"<f " in buf or b"<f>" in buf or b"<f/" in buf):
                    has_formula = True
                if not has_embed_img and b"DISPIMG" in buf:
                    has_embed_img = True
                for m in HROW_RE.finditer(buf):
                    for k, v in ATTR_NUM_RE.findall(m.group(0)):
                        if k == b"r":
                            r = int(v)
                            if r not in hidden_rows:
                                hidden_rows.append(r)
                for m in HCOL_RE.finditer(buf):
                    attrs = dict(ATTR_NUM_RE.findall(m.group(0)))
                    lo = int(attrs.get(b"min", b"1"))
                    hi = int(attrs.get(b"max", attrs.get(b"min", b"1")))
                    for c in range(lo, min(hi, 256) + 1):
                        if c not in hidden_cols:
                            hidden_cols.append(c)
                if part == 1:
                    m = DIM_RE.search(buf)
                    if m:
                        rng = parse_range(m.group(1).decode("ascii", "replace"))
                        if rng:
                            dim_rows, dim_cols = rng[2], rng[3]
                for m in ROW_RE.finditer(buf):
                    v = int(m.group(1))
                    if v > max_row:
                        max_row = v
                if part <= 3:
                    for m in COL_RE.finditer(buf):
                        n = 0
                        for ch in m.group(1):
                            n = n * 26 + (ch - 64)
                        if n > max_col:
                            max_col = n
                if do_merge:
                    m = MERGE_BLOCK_RE.search(buf)
                    if m:
                        for ref in MERGE_ONE_RE.findall(m.group(1)):
                            rng = parse_range(ref.decode("ascii", "replace"))
                            if rng:
                                merges.append(rng)
                                max_col = max(max_col, rng[3])
                        break
                    tail = buf[-8192:]
        max_row = max(max_row, dim_rows)
        max_col = max(max_col, dim_cols)
    return (merges, has_formula, max_row, max_col, names_in_wb,
            hidden_cols, hidden_rows, has_embed_img)


# 与宿主（Rust）约定的退出码，需要随 #FATAL 一并回报
EXIT_ERROR = 1   # 运行期错误
EXIT_USAGE = 3   # 参数/引用错误（文件不存在、工作表不存在、行号越界……）


class HelperError(Exception):
    def __init__(self, msg, code=EXIT_ERROR):
        super().__init__(msg)
        self.code = code


def list_sheets(path):
    kind = sniff_format(path)
    if kind == "xls":
        import xlrd
        book = xlrd.open_workbook(path)
        return list(book.sheet_names())
    if kind == "xlsx":
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(path) as z:
            wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        return [sh.get("name") for sh in wb_xml.iter(NS_MAIN + "sheet")]
    raise HelperError("无法识别的文件格式")


# --------------------------------------------------------------------------
# mode = probe
# --------------------------------------------------------------------------
def _flat_pair(c, v):
    """#HFILL 的一对（列号, 值）：值已转义，不会有真 tab，宿主按两项一组解析。"""
    return "%d\t%s" % (int(c), esc(v))


def header_fill(path, sheet, header_row, merges, kind, max_col):
    """字段名行落在纵向合并区里时，字段名在合并区左上角那一行。

    两行表头的常见写法（D1:D2 合并）会让字段名行读到空值，
    这里回读合并区上方的行，把值取出来交给宿主。
    返回 [(列号(1起), 值), ...]
    """
    rows = sorted({r1 for (r1, c1, r2, c2) in merges if r1 < header_row <= r2})
    if not rows or header_row <= 1:
        return []
    lo, hi = rows[0], min(rows[-1], header_row - 1)
    fetched = {}
    try:
        if kind == "xls":
            import xlrd
            book = xlrd.open_workbook(path)
            reader = XlsReader(book, book.sheet_by_name(sheet))
        else:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            reader = XlsxReader(wb[sheet])
        for r, cells in reader.iter_rows(lo, hi, max_col):
            fetched[r] = cells
    except Exception:
        return []
    out = []
    for (r1, c1, r2, c2) in merges:
        if not (r1 < header_row <= r2):
            continue
        row = fetched.get(r1)
        if not row:
            continue
        ci = c1 - 1
        if 0 <= ci < len(row) and row[ci][0] != K_EMPTY:
            out.append((c1, row[ci][1]))
    return out


def cmd_probe(path, sheet, header_row=0):
    header_row = int(header_row)
    if not os.path.exists(path):
        raise HelperError("文件不存在: %s" % path, EXIT_USAGE)
    kind = sniff_format(path)
    if kind == "xls":
        import xlrd
        try:
            book = xlrd.open_workbook(path, formatting_info=True)
        except Exception:
            try:
                book = xlrd.open_workbook(path)
            except Exception as e:
                raise HelperError("打开 .xls 失败（可能已加密或损坏）: %s" % e)
        names = list(book.sheet_names())
        if sheet not in names:
            raise HelperError("工作表 %r 不存在。可用工作表: %s" % (sheet, "、".join(names)),
                              EXIT_USAGE)
        sh = book.sheet_by_name(sheet)
        merges = []
        for rlo, rhi, clo, chi in getattr(sh, "merged_cells", []) or []:
            merges.append((rlo + 1, clo + 1, rhi, chi))    # xlrd 为开区间 → 闭区间
        n_rows, n_cols = int(sh.nrows), int(sh.ncols)
        if merges:
            n_cols = max(n_cols, max(m[3] for m in merges))
        hidden_cols, hidden_rows = [], []
        try:
            hidden_cols = sorted(c + 1 for c, info in (getattr(sh, "colinfo_map", None) or {}).items()
                                 if getattr(info, "hidden", 0))
            hidden_rows = sorted(r + 1 for r, info in (getattr(sh, "rowinfo_map", None) or {}).items()
                                 if getattr(info, "hidden", 0))
        except Exception:
            pass
        w("#KIND\tExcel 2003 (.xls)")
        w("#SHEET\t%s\t%d\t%d" % (esc(sheet), n_rows, n_cols))
        w("#SHEETS\t%s" % TAB.join(esc(x) for x in names))
        w("#MERGES\t%s" % "|".join("%d,%d,%d,%d" % m for m in merges))
        w("#FORMULA\t0")
        w("#HIDDEN\t%s\t%s" % (",".join(str(x) for x in hidden_cols),
                               ",".join(str(x) for x in hidden_rows)))
        w("#EMBEDIMG\t0")
        hf = header_fill(path, sheet, header_row, merges, "xls", n_cols)
        w("#HFILL\t" + TAB.join(_flat_pair(c, v) for c, v in hf))
        return 0

    if kind == "xlsx":
        (merges, has_formula, max_row, max_col, names,
         hidden_cols, hidden_rows, has_img) = xlsx_scan(path, sheet)
        w("#KIND\tExcel 2007+ (.xlsx)")
        w("#SHEET\t%s\t%d\t%d" % (esc(sheet), max_row, max_col))
        w("#SHEETS\t%s" % TAB.join(esc(x) for x in names))
        w("#MERGES\t%s" % "|".join("%d,%d,%d,%d" % m for m in merges))
        w("#FORMULA\t%d" % (1 if has_formula else 0))
        w("#HIDDEN\t%s\t%s" % (",".join(str(x) for x in hidden_cols),
                               ",".join(str(x) for x in hidden_rows)))
        w("#EMBEDIMG\t%d" % (1 if has_img else 0))
        hf = header_fill(path, sheet, header_row, merges, "xlsx", max_col)
        w("#HFILL\t" + TAB.join(_flat_pair(c, v) for c, v in hf))
        return 0

    raise HelperError("无法识别的文件格式（既不是 .xlsx 也不是 .xls），请检查文件是否完整")


# --------------------------------------------------------------------------
# mode = dump
# --------------------------------------------------------------------------
def cmd_dump(path, sheet, header_row, data_row, last_row, step, offset, hints, n_cols,
             warm=0, max_read=0):
    kind = sniff_format(path)
    date_hints = None
    reader = None

    if kind == "xlsx":
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
        except Exception as e:
            raise HelperError("打开 .xlsx 失败: %s" % e)
        if sheet not in wb.sheetnames:
            raise HelperError("工作表 %r 不存在。可用工作表: %s"
                              % (sheet, "、".join(wb.sheetnames)))
        ws = wb[sheet]
        reader = XlsxReader(ws)
    elif kind == "xls":
        import xlrd
        try:
            book = xlrd.open_workbook(path)
        except Exception as e:
            raise HelperError("打开 .xls 失败: %s" % e)
        if sheet not in book.sheet_names():
            raise HelperError("工作表 %r 不存在。可用工作表: %s"
                              % (sheet, "、".join(book.sheet_names())), EXIT_USAGE)
        reader = XlsReader(book, book.sheet_by_name(sheet))
    else:
        raise HelperError("无法识别的文件格式")

    # 列数由宿主（probe 阶段的真实扫描结果）给出：流式写出的 xlsx 其
    # <dimension> 常常只写 A1，不能信 ws.max_column。
    n_cols = max(int(n_cols), 1)
    last = reader.n_rows if last_row <= 0 else last_row
    if last < header_row:
        raise HelperError("第 %d 行（字段名称所在行）超出工作表范围（共 %d 行）"
                          % (header_row, reader.n_rows), EXIT_USAGE)

    total = max(last - header_row + 1, 1)
    tick = max(total // 100, 1)
    read = 0
    last_seen = header_row - 1
    emitted = 0

    for r, cells in reader.iter_rows(header_row, last, n_cols):
        if max_read and read >= max_read:
            break
        read += 1
        last_seen = r
        if read % tick == 0 or read == total:
            progress(read, total)

        if r == header_row:
            if date_hints is None:
                names = [c[1] for c in cells]
                date_hints = [bool(hints) and any(
                    h in (nm or "").lower() for h in DATE_NAME_HINTS) for nm in names]
            w("#HEADER\t" + TAB.join(encode(c) for c in cells))
            continue

        if r < data_row:
            continue
        # 开头 warm 行无条件判定，其余按步长跳跃抽样
        if step > 1 and (r - data_row) >= warm and (r - offset) % step != 0:
            continue

        if date_hints is None:
            date_hints = [False] * n_cols
        for i in range(len(cells)):
            if i >= len(date_hints):
                break
            if date_hints[i] and cells[i][0] in (K_TEXT, K_INT, K_DEC) and cells[i][1]:
                d = parse_date_like(cells[i][1], allow_compact=True)
                if d:
                    cells[i] = cell_from_date_kind(d[0], d[1])
        w("#ROW\t%d\t%s" % (r, TAB.join(encode(c) for c in cells)))
        emitted += 1

    progress(total, total)
    pct = getattr(reader, "pct_cells", 0)
    if pct:
        w("#PCT\t%d" % pct)
    tz = getattr(reader, "tz_dropped", None) or {}
    if tz:
        w("#TZ\t" + TAB.join("%d\t%d" % (c, tz[c]) for c in sorted(tz)))
    w("#END\t%d\t%d\t%d" % (read, last_seen, emitted))
    return 0


class XlsxReader:
    """只读模式单次顺序遍历：整个工作表只解析一遍 XML，绝不按行重开流。"""

    def __init__(self, ws):
        self.ws = ws
        self.n_rows = ws.max_row or 0
        self.n_cols = ws.max_column or 0
        self.pct_cells = 0       # 带百分号格式的数值单元格
        self.tz_dropped = {}     # 列号(1起) -> 带时区偏移、偏移被丢弃的单元格数

    def iter_rows(self, first, last, n_cols):
        n_cols = max(n_cols, 1)
        r = first
        for row in self.ws.iter_rows(min_row=first, max_row=last,
                                     min_col=1, max_col=n_cols):
            if not row:
                continue
            r = getattr(row[0], "row", None) or r
            while len(row) < n_cols:
                row = tuple(row) + (None,)
            cells = []
            for ci, c in enumerate(row[:n_cols]):
                v = getattr(c, "value", None) if c is not None else None
                # 百分比格式：Excel 里显示 12.5%，实际存储值是 0.125
                if isinstance(v, (int, float)) and not isinstance(v, bool) \
                        and "%" in (getattr(c, "number_format", "") or ""):
                    self.pct_cells += 1
                t = cell_from_openpyxl(c)
                # 带时区偏移的日期时间：MySQL 的 datetime 存不下时区，偏移会被丢掉
                if t[0] in (K_DATE, K_DATETIME, K_TIME) and isinstance(v, str) \
                        and RE_TZ_OFFSET.search(v.strip()):
                    self.tz_dropped[ci + 1] = self.tz_dropped.get(ci + 1, 0) + 1
                cells.append(t)
            yield r, cells
            r += 1


def cell_from_openpyxl(c):
    v = getattr(c, "value", None) if c is not None else None
    if v is None:
        return CELL_EMPTY
    if isinstance(v, bool):
        return (K_BOOL, "1" if v else "0")
    if isinstance(v, (dt.datetime, dt.date, dt.time, dt.timedelta)):
        return classify_native_date(v, getattr(c, "number_format", None))
    if isinstance(v, (int, float)):
        if isinstance(v, float):
            if math.isnan(v) or math.isinf(v):
                return (K_ERROR, "")
            if v.is_integer() and abs(v) < 1e15:
                return (K_INT, str(int(v)))
            return (K_DEC, repr(v))
        return (K_INT, str(v))
    if isinstance(v, str):
        if getattr(c, "data_type", None) == "e":
            return (K_ERROR, clean(v))
        return classify_string(v)
    return (K_TEXT, clean(v))


class XlsReader:
    def __init__(self, book, sh):
        self.XL = __import__("xlrd")
        self.book = book
        self.sh = sh
        self.n_rows = int(sh.nrows)
        self.n_cols = int(sh.ncols)
        self.pct_cells = 0
        self.tz_dropped = {}     # 列号(1起) -> 带时区偏移、偏移被丢弃的单元格数

    def iter_rows(self, first, last, n_cols):
        for r in range(first, last + 1):
            yield r, self.row(r, max(n_cols, 1))

    def row(self, r, n_cols):
        XL = self.XL
        datemode = self.book.datemode
        sh = self.sh
        ri = r - 1
        out = []
        for ci in range(n_cols):
            if ci >= sh.ncols or ri >= sh.nrows:
                out.append(CELL_EMPTY)
                continue
            ctype = sh.cell_type(ri, ci)
            value = sh.cell_value(ri, ci)
            if ctype in (XL.XL_CELL_EMPTY, XL.XL_CELL_BLANK):
                out.append(CELL_EMPTY)
            elif ctype == XL.XL_CELL_TEXT:
                t = classify_string(str(value))
                # 带时区偏移的日期时间：MySQL 的 datetime 存不下时区，偏移会被丢掉
                if t[0] in (K_DATE, K_DATETIME, K_TIME) and RE_TZ_OFFSET.search(str(value).strip()):
                    self.tz_dropped[ci + 1] = self.tz_dropped.get(ci + 1, 0) + 1
                out.append(t)
            elif ctype == XL.XL_CELL_BOOLEAN:
                out.append((K_BOOL, "1" if value else "0"))
            elif ctype == XL.XL_CELL_ERROR:
                out.append((K_ERROR, ""))
            elif ctype == XL.XL_CELL_DATE:
                try:
                    d = XL.xldate_as_datetime(value, datemode)
                except Exception:
                    out.append((K_TEXT, clean(value)))
                    continue
                if d.year <= 1900:
                    out.append((K_TIME, fmt_time(d.time())))
                elif d.hour == d.minute == d.second == 0:
                    out.append((K_DATE, d.strftime("%Y-%m-%d")))
                else:
                    out.append((K_DATETIME, d.strftime("%Y-%m-%d %H:%M:%S")))
            else:
                if isinstance(value, float) and value.is_integer() and abs(value) < 1e15:
                    out.append((K_INT, str(int(value))))
                else:
                    out.append((K_DEC, repr(float(value))))
        return out


# --------------------------------------------------------------------------
# 极简 xlsx 写出：只用标准库 zipfile 拼 XML，不依赖 openpyxl，也不产生任何
# 临时文件。openpyxl 保存时会建中转文件再删掉，在 /tmp 很小（10MB tmpfs）或
# 进程不允许删文件的受限环境里会直接失败；errrows.xlsx 结构简单，自己写更稳。
# --------------------------------------------------------------------------
_XLSX_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"""

_XLSX_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""

_XLSX_WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""

_XLSX_WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>"""

# XML 1.0 不允许这些控制字符，Excel 也会拒绝加载
_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")
_CELL_MAX = 32767          # Excel 单元个字符上限
_SHEET_MAX = 31            # 工作表名长度上限


def _xml_text(v):
    s = _XML_BAD.sub("", str(v))
    if len(s) > _CELL_MAX:
        s = s[:_CELL_MAX]
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _xml_attr(v):
    return _xml_text(v).replace('"', "&quot;")


def col_name(i):
    """1 -> A, 27 -> AA"""
    name = ""
    while i > 0:
        i, r = divmod(i - 1, 26)
        name = chr(65 + r) + name
    return name


def _num_text(v):
    if isinstance(v, int):
        return str(v)
    return repr(float(v))


def _cell_xml(ref, v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return '<c r="%s" t="b"><v>%d</v></c>' % (ref, 1 if v else 0)
    if isinstance(v, (int, float)):
        return '<c r="%s"><v>%s</v></c>' % (ref, _num_text(v))
    if v == "":
        return ""
    return '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
        ref, _xml_text(v))


def write_xlsx(path, rows, sheet="Sheet1", dims=None):
    """把 rows（可迭代，每项是一行的 list/tuple）写成一个极简 xlsx。

    逐行流式写入，内存占用与行数无关；不建临时文件、不删除任何文件。
    dims=(总行数, 总列数) 可选，给了就写出 <dimension>，方便其它工具直接取范围
    （只读模式下 openpyxl 拿得到 max_row/max_column，不用扫全表）。
    """
    name = _xml_attr((sheet or "Sheet1")[:_SHEET_MAX]) or "Sheet1"
    dim = ""
    if dims and dims[0] > 0 and dims[1] > 0:
        dim = '<dimension ref="A1:%s%d"/>' % (col_name(dims[1]), dims[0])
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _XLSX_CONTENT_TYPES)
        z.writestr("_rels/.rels", _XLSX_ROOT_RELS)
        z.writestr("xl/workbook.xml", _XLSX_WORKBOOK % name)
        z.writestr("xl/_rels/workbook.xml.rels", _XLSX_WB_RELS)
        with z.open("xl/worksheets/sheet1.xml", "w") as f:
            f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    b'<worksheet xmlns="http://schemas.openxmlformats.org/'
                    b'spreadsheetml/2006/main">')
            if dim:
                f.write(dim.encode("utf-8"))
            f.write(b"<sheetData>")
            for ri, row in enumerate(rows, 1):
                f.write(('<row r="%d">' % ri).encode("utf-8"))
                for ci, v in enumerate(row, 1):
                    xml = _cell_xml("%s%d" % (col_name(ci), ri), v)
                    if xml:
                        f.write(xml.encode("utf-8"))
                f.write(b"</row>")
            f.write(b"</sheetData></worksheet>")


# --------------------------------------------------------------------------
# mode = errrows   从 stdin 读失败行，写出 xlsx
# --------------------------------------------------------------------------
def cmd_errrows(out_path):
    header = []
    rows = []
    for ln in sys.stdin.read().split("\n"):
        if not ln:
            continue
        parts = ln.split(TAB)
        tag = parts[0]
        if tag == "H":
            header = [unesc(x) for x in parts[1:]]
        elif tag == "R":
            vals = [unesc(x) for x in parts[1:]]
            # 第一列「原行号」写成数值，方便在 Excel 里排序/筛选
            if vals:
                try:
                    vals[0] = int(str(vals[0]).strip())
                except (TypeError, ValueError):
                    pass
            rows.append(vals)

    def gen():
        yield header
        for r in rows:
            yield r

    n_cols = max([len(header)] + [len(r) for r in rows]) if header or rows else 0
    write_xlsx(out_path, gen(), "errrows", dims=(1 + len(rows), n_cols))
    w("#OK\t%s" % os.path.abspath(out_path))
    return 0


# --------------------------------------------------------------------------
# 入口
# --------------------------------------------------------------------------
def main(argv):
    if not argv:
        note("helper: 缺少 mode 参数")
        return 2
    mode = argv[0]
    try:
        if mode == "probe":
            return cmd_probe(argv[1], argv[2],
                             int(argv[3]) if len(argv) > 3 else 0)
        if mode == "dump":
            return cmd_dump(argv[1], argv[2], int(argv[3]), int(argv[4]),
                            int(argv[5]), int(argv[6]), int(argv[7]),
                            argv[8] == "1", int(argv[9]),
                            int(argv[10]) if len(argv) > 10 else 0,
                            int(argv[11]) if len(argv) > 11 else 0)
        if mode == "errrows":
            return cmd_errrows(argv[1])
        if mode == "list":
            w("#SHEETS\t%s" % TAB.join(esc(x) for x in list_sheets(argv[1])))
            return 0
        note("helper: 未知 mode %r" % mode)
        return 2
    except HelperError as e:
        w("#FATAL\t%d\t%s" % (e.code, esc(str(e))))
        return e.code
    except ImportError as e:
        w("#FATAL\t%d\t%s" % (EXIT_ERROR,
                               esc("缺少 Python 依赖: %s。请先执行 pip install openpyxl xlrd" % e)))
        return EXIT_ERROR
    except BrokenPipeError:
        return 0
    except Exception as e:
        w("#FATAL\t%d\t%s" % (EXIT_ERROR, esc("%s: %s" % (type(e).__name__, e))))
        return EXIT_ERROR


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", newline="\n")
        sys.stderr.reconfigure(encoding="utf-8", newline="\n")
    except Exception:
        pass
    sys.exit(main(sys.argv[1:]))"##;

// ===========================================================================
// 常量
// ===========================================================================
const VERSION: &str = "2.2.0";

const EXIT_OK: i32 = 0;
const EXIT_ERROR: i32 = 1;
const EXIT_FORMAT: i32 = 2;
const EXIT_USAGE: i32 = 3;

const INT_RANGES: &[(&str, i128, i128)] = &[
    ("tinyint", -128, 127),
    ("smallint", -32768, 32767),
    ("mediumint", -8388608, 8388607),
    ("int", -2147483648, 2147483647),
    ("bigint", -9223372036854775808, 9223372036854775807),
];

const VARCHAR_LADDER: &[i64] = &[16, 32, 50, 64, 100, 128, 150, 192, 255, 320, 400, 512, 768, 1000];

const ID_HINTS: &[&str] = &[
    "手机", "电话", "座机", "联系电话", "联系方式", "联系", "号码", "身份证", "证件号",
    "银行卡", "卡号", "账号", "帐号", "工号", "学号", "邮箱", "邮编", "邮政编码",
    "phone", "mobile", "telephone", "email", "idcard", "id_card", "passport",
    "zipcode", "zip_code",
];

const MONEY_HINTS: &[&str] = &[
    "金额", "价格", "单价", "总价", "费用", "运费", "工资", "薪水", "薪资",
    "税额", "金额合计", "amount", "price", "money", "cost", "fee", "salary", "wage",
];

const RESERVED_HINTS: &[&str] = &[
    "table", "select", "order", "group", "key", "index", "from", "where",
];

const SUBTOTAL_HINTS: &[&str] = &[
    "合计", "小计", "总计", "汇总", "共计", "total", "subtotal", "summary",
];

// ===========================================================================
// 错误
// ===========================================================================
struct CliError {
    msg: String,
    code: i32,
}

impl CliError {
    fn new(msg: impl Into<String>, code: i32) -> Self {
        CliError { msg: tr(&msg.into()), code }
    }
    fn usage(msg: impl Into<String>) -> Self {
        CliError::new(msg, EXIT_USAGE)
    }
    fn general(msg: impl Into<String>) -> Self {
        CliError::new(msg, EXIT_ERROR)
    }
}

// ===========================================================================
// 小工具
// ===========================================================================
fn human_num(n: i64) -> String {
    let neg = n < 0;
    let s = n.unsigned_abs().to_string();
    let len = s.len();
    let mut out = String::new();
    for (i, ch) in s.chars().enumerate() {
        if i > 0 && (len - i) % 3 == 0 {
            out.push(',');
        }
        out.push(ch);
    }
    if neg { format!("-{}", out) } else { out }
}

fn fmt_dur(sec: f64) -> String {
    // 时长文案带数字宽度格式（{:.0}/{:.1}），不走翻译表，直接按语言分支拼
    let zh = lang_zh();
    if sec < 1.0 {
        let ms = sec * 1000.0;
        if zh { format!("{:.0} 毫秒", ms) } else { format!("{:.0} ms", ms) }
    } else if sec < 60.0 {
        if zh { format!("{:.1} 秒", sec) } else { format!("{:.1} s", sec) }
    } else if sec < 3600.0 {
        let m = (sec / 60.0).floor();
        if zh { format!("{:.0} 分 {:.0} 秒", m, sec - m * 60.0) }
        else { format!("{:.0}m {:.0}s", m, sec - m * 60.0) }
    } else {
        let h = (sec / 3600.0).floor();
        let m = ((sec - h * 3600.0) / 60.0).floor();
        if zh { format!("{:.0} 时 {:.0} 分", h, m) } else { format!("{:.0}h {:.0}m", h, m) }
    }
}

fn char_w(c: char) -> usize {
    let cp = c as u32;
    if cp == 0 || cp < 0x20 || (0x7f..0xa0).contains(&cp) {
        return 0;
    }
    let wide = (0x1100..=0x115f).contains(&cp)
        || (0x2e80..=0x303e).contains(&cp)
        || (0x3041..=0x33ff).contains(&cp)
        || (0x3400..=0x4dbf).contains(&cp)
        || (0x4e00..=0x9fff).contains(&cp)
        || (0xa000..=0xa4cf).contains(&cp)
        || (0xac00..=0xd7a3).contains(&cp)
        || (0xf900..=0xfaff).contains(&cp)
        || (0xfe30..=0xfe6f).contains(&cp)
        || (0xff00..=0xff60).contains(&cp)
        || (0xffe0..=0xffe6).contains(&cp)
        || (0x20000..=0x3fffd).contains(&cp);
    if wide { 2 } else { 1 }
}

fn disp_width(s: &str) -> usize {
    s.chars().map(char_w).sum()
}

fn pad_to(s: &str, width: usize, right: bool) -> String {
    let w = disp_width(s);
    if w >= width {
        return s.to_string();
    }
    let fill = " ".repeat(width - w);
    if right { format!("{}{}", fill, s) } else { format!("{}{}", s, fill) }
}

fn one_line(s: &str) -> String {
    s.replace('\n', " ").replace('\r', " ").replace('\t', " ")
}

fn truncate_disp(s: &str, width: usize) -> String {
    if disp_width(s) <= width {
        return s.to_string();
    }
    // 省略号跟着语言走：中文用 …，英文用 ...（英文里三个点更符合习惯）
    let ell = if lang_zh() { "…" } else { "..." };
    let mut out = String::new();
    let mut w = 0usize;
    let limit = width.saturating_sub(disp_width(ell));
    for ch in s.chars() {
        let cw = char_w(ch);
        if w + cw > limit {
            break;
        }
        out.push(ch);
        w += cw;
    }
    out.push_str(ell);
    out
}

/// 画一张带边框的表格（按显示宽度对齐，中日韩宽字符算 2 列）。
/// 表头与单元格先翻译再算宽度 —— 中文和英文的显示宽度差很多，顺序反了框就歪了。
fn render_table(headers: &[&str], rows: &[Vec<String>]) -> String {
    let ncol = headers.len();
    let hs: Vec<String> = headers.iter().map(|h| tr(h)).collect();
    let mut widths: Vec<usize> = hs.iter().map(|h| disp_width(h)).collect();
    // 英文句子比中文长，单元格截断宽度也放宽一点，否则信息会被砍掉
    let cell_cap = if lang_zh() { 60 } else { 76 };
    let mut clean: Vec<Vec<String>> = Vec::with_capacity(rows.len());
    for r in rows {
        let mut row = Vec::with_capacity(ncol);
        for i in 0..ncol {
            let mut v = one_line(&tr(r.get(i).map(|s| s.as_str()).unwrap_or("")));
            if disp_width(&v) > cell_cap {
                v = truncate_disp(&v, cell_cap);
            }
            let w = disp_width(&v);
            if w > widths[i] {
                widths[i] = w;
            }
            row.push(v);
        }
        clean.push(row);
    }
    let rule = |l: &str, m: &str, r: &str| -> String {
        let mut s = String::from(l);
        for (i, w) in widths.iter().enumerate() {
            s.push_str(&"─".repeat(w + 2));
            s.push_str(if i + 1 == ncol { r } else { m });
        }
        s
    };
    let row_line = |cells: &[String]| -> String {
        let mut s = String::from("│");
        for (i, v) in cells.iter().enumerate() {
            s.push(' ');
            s.push_str(&pad_to(v, widths[i], false));
            s.push(' ');
            s.push('│');
        }
        s
    };
    let mut out = String::new();
    out.push_str(&rule("┌", "┬", "┐"));
    out.push('\n');
    out.push_str(&row_line(&hs));
    out.push('\n');
    out.push_str(&rule("├", "┼", "┤"));
    for row in &clean {
        out.push('\n');
        out.push_str(&row_line(row));
    }
    out.push('\n');
    out.push_str(&rule("└", "┴", "┘"));
    out
}


// ===========================================================================
// 语言与文案（i18n）
//   * 中文环境输出中文，其它环境输出英文；
//   * 判定顺序：--lang=xx > XLSXTOMYSQL_LANG > LC_ALL > LC_MESSAGES > LANG；
//     全都没有时：Unix 按 POSIX 习惯当英文，Windows 读系统 UI 语言；
//   * 所有界面文案收在 CATALOG，以中文原文为键。键里的 {} 是运行期填入的值，
//     译文必须与中文的 {} 个数、顺序一致；
//   * 接入点分两类：
//       - 明确知道是文案的地方（标语、表头、错误构造）用 t() / tf()；
//       - 输出边界（Style::wrap / Out::line / render_table 单元格）用 tr()，
//         先整串精确匹配，再按模板做"整串锚定"匹配 —— 锚定保证不会误伤表格里的数据。
// ===========================================================================
static LANG_ZH: AtomicBool = AtomicBool::new(true);

fn lang_zh() -> bool {
    LANG_ZH.load(Ordering::Relaxed)
}

/// 设定语言；mode 为空时按环境变量自动判断，否则 zh* → 中文、en* → 英文、其它按自动。
fn set_lang(mode: &str) {
    let m = mode.trim().to_lowercase();
    let zh = if m.is_empty() || m == "auto" {
        detect_lang_from_env()
    } else if m.starts_with("zh") || m == "cn" {
        true
    } else if m.starts_with("en") {
        false
    } else {
        detect_lang_from_env()
    };
    LANG_ZH.store(zh, Ordering::Relaxed);
}

fn detect_lang_from_env() -> bool {
    for k in ["XLSXTOMYSQL_LANG", "LC_ALL", "LC_MESSAGES", "LANG"] {
        if let Ok(v) = std::env::var(k) {
            let v = v.trim().to_lowercase();
            if !v.is_empty() {
                return v.starts_with("zh");
            }
        }
    }
    default_lang()
}

#[cfg(windows)]
fn default_lang() -> bool {
    // Windows 常常没有 LC_* 变量，读系统 UI 语言：主语言 ID 0x04 = 中文
    #[link(name = "kernel32")]
    extern "system" {
        fn GetUserDefaultUILanguage() -> u16;
    }
    let langid = unsafe { GetUserDefaultUILanguage() };
    (langid & 0x3ff) == 0x04
}

#[cfg(not(windows))]
fn default_lang() -> bool {
    // POSIX 下没有任何 locale 变量等价于 C locale，按英文处理
    false
}

/// 是否需要查译文表。判断依据是"含汉字或中文标点"：
///   CJK 标点(3000-303F) / 汉字(4E00-9FFF) / 兼容汉字(F900-FAFF) / 全角(FF00-FFEF)
///   外加通用标点(2000-206F，含 —— … “ ”) 与箭头(2190-21FF，含 →)。
/// 不含这些字符的串一定不在译文表里，直接跳过匹配 —— 表格里绝大多数单元格是数据，
/// 这个判断省掉大量无用扫描。制表符(2500-257F)故意不算，表格框线不该触发查表。
fn needs_translation(s: &str) -> bool {
    s.chars().any(|c| matches!(c,
        '\u{2000}'..='\u{206f}' | '\u{2190}'..='\u{21ff}' | '\u{3000}'..='\u{303f}'
        | '\u{4e00}'..='\u{9fff}' | '\u{f900}'..='\u{faff}' | '\u{ff00}'..='\u{ffef}'))
}

// ==== I18N CATALOG BEGIN (由 tools/apply_i18n.py 生成，勿手改) ====
/// 译文表：中文原文 → 英文。键里 {} 是运行期填入的值。
/// 生成自 tools/i18n_catalog.py（改文案请改那里再跑 tools/apply_i18n.py --write）。
const CATALOG: &[(&str, &str)] = &[
    // ---- 整串精确匹配 ----
    ("  ⚠ 含公式；当前按「计算结果」读取，若整列为空请先用 Excel 打开保存一次", "  ! formulas detected; cached results are read, so open and save the file once in Excel if a column comes out empty"),
    ("  ⚠ 检测到「置于单元格内」的图片（WPS 的 DISPIMG / Excel 单元格图片）；图片本身无法写入 SQL，对应单元格会变成空值", "  ! images placed inside cells detected (WPS DISPIMG / Excel in-cell pictures); images cannot be written to SQL, so those cells become empty"),
    ("  数据区的合并单元格会按左上角内容自动补齐整块（--no-fill-down 可关闭）", "  merged cells in the data area are filled from the top-left value (disable with --no-fill-down)"),
    ("(Excel 错误值)", "(Excel error)"),
    ("(空)", "(empty)"),
    ("--empty-as-null 只能是 auto / always / never", "--empty-as-null must be auto / always / never"),
    ("--merge-scan 只能是 auto / on / off", "--merge-scan must be auto / on / off"),
    ("--sample-ratio 需要小数", "--sample-ratio expects a decimal number"),
    ("Excel 错误值，已按空处理", "Excel error cell, treated as empty"),
    ("MySQL 类型", "MySQL type"),
    ("helper: 缺少 mode 参数", "helper: missing the mode argument"),
    ("stderr 已设为管道", "stderr pipe missing (internal error)"),
    ("stdout 已设为管道", "stdout pipe missing (internal error)"),
    ("…", "..."),
    ("「共几行」不能是负数", "ROW_COUNT cannot be negative"),
    ("「共几行」需要是非负整数", "ROW_COUNT must be a non-negative integer"),
    ("「字段」", "[field]"),
    ("「字段名称所在行」从 1 开始，不能小于 1", "HEADER_ROW is 1-based and cannot be less than 1"),
    ("「字段名称所在行」需要是正整数", "HEADER_ROW must be a positive integer"),
    ("「第一个数据所在行」从 1 开始，不能小于 1", "FIRST_DATA_ROW is 1-based and cannot be less than 1"),
    ("「第一个数据所在行」需要是正整数", "FIRST_DATA_ROW must be a positive integer"),
    ("」", "]"),
    ("【1/2】预扫描 —— 识别字段类型 & 检查表格格式", "[1/2] Pre-scan - detect column types and check the table layout"),
    ("【2/2】转换 —— 生成 SQL", "[2/2] Convert - writing SQL"),
    ("不是有效的 .xlsx 文件（缺少工作簿定义）", "not a valid .xlsx file (no workbook part found)"),
    ("与 SQL 关键字同名(已用反引号包裹)", "same name as an SQL keyword (back-quoted)"),
    ("仅扫描模式，未生成 sql。", "scan-only mode: no SQL was written."),
    ("从「第一个数据所在行」开始没有任何数据，请检查行号参数", "no data from FIRST_DATA_ROW onwards; check the row number"),
    ("以数字开头", "starts with a digit"),
    ("位置", "Where"),
    ("值", "Value"),
    ("全部为整数", "all integers"),
    ("全部为日期", "dates only"),
    ("全部为日期时间", "datetimes only"),
    ("全部为时间", "times only"),
    ("共多少行数据", "Data rows"),
    ("内容超过 text 上限(65535 字节)，建议改用 longtext", "content exceeds the text limit (65535 bytes); use longtext instead"),
    ("分隔用的空行", "blank separator row"),
    ("分隔线/装饰行", "divider / decoration row"),
    ("原字段名为空", "original name is empty"),
    ("原字段名全为非法字符", "original name is entirely illegal characters"),
    ("原行号", "Row"),
    ("只有真假值", "boolean values only"),
    ("含空白", "contains whitespace"),
    ("含非法字符", "contains illegal characters"),
    ("失败原因", "Reason"),
    ("失败原因归类", "Failure reasons"),
    ("失败行数", "Failed"),
    ("失败行文件", "Failed-rows file"),
    ("字段名为空", "empty column name"),
    ("字段名处理", "Column-name notes"),
    ("字段名所在行", "Header row"),
    ("工作表", "Sheet"),
    ("工作表  :", "sheet  :"),
    ("已丢弃时间部分", "the time part was dropped"),
    ("已读过表头", "header already read"),
    ("成功行数", "Succeeded"),
    ("扫描中", "scanning"),
    ("扫描方式:", "scan mode:"),
    ("按 Excel 日期序列号转换", "converted from the Excel date serial number"),
    ("推断依据", "Basis"),
    ("提醒", "Notice"),
    ("数据区中间出现整行为空的行，会打断数据；请删除该行或加 --force 忽略", "a completely blank row appears in the middle of the data and breaks the sequence; delete it, or add --force to ignore it"),
    ("数据区为空", "empty data area"),
    ("数据超出字段名范围", "data beyond the header range"),
    ("文件格式", "Format"),
    ("文件格式:", "format:"),
    ("新字段名", "New column"),
    ("无法识别的文件格式", "unrecognised file format"),
    ("无法识别的文件格式（既不是 .xlsx 也不是 .xls），请检查文件是否完整", "unrecognised file format (neither .xlsx nor .xls); check that the file is intact"),
    ("日期与日期时间混合，统一为 datetime", "dates and datetimes mixed, unified as datetime"),
    ("日期与时间混杂，按文本保存以免失真", "dates and times mixed, stored as text to avoid corruption"),
    ("未能读到字段名行，请检查「字段名称所在行」参数是否正确", "could not read the header row; check the HEADER_ROW argument"),
    ("末尾空行", "trailing blank rows"),
    ("样本含科学计数法数值，使用 double", "the sample contains scientific-notation values, using double"),
    ("检测到 .xlsb 二进制工作簿，openpyxl 不支持，请另存为 .xlsx 后重试", "this is an .xlsb binary workbook, which openpyxl cannot read; save it as .xlsx and retry"),
    ("源文件", "Source file"),
    ("疑似合计/小计行", "looks like a total / subtotal row"),
    ("疑似说明行", "looks like a note row"),
    ("第一条数据所在行", "First data row"),
    ("类型扫描方式", "Scan mode"),
    ("结果表 1 · 转换结果", "Report 1 · Conversion result"),
    ("结果表 2 · 新字段名称及类型", "Report 2 · New column names and types"),
    ("耗时", "Elapsed"),
    ("行", "rows"),
    ("行数", "Rows"),
    ("警告行数", "Warnings"),
    ("该行内容与字段名行相同，疑似重复表头；请删除后重试", "this row repeats the header row; delete it and retry"),
    ("说明", "Details"),
    ("请按上面的位置提示处理源文件后重新执行；确需强行转换可追加 --force", "fix the source file as indicated above and run again; add --force to convert anyway"),
    ("超长截断", "truncated (too long)"),
    ("跳过空行", "Blank rows skipped"),
    ("转换中", "converting"),
    ("输出文件", "Output file"),
    ("重复的字段名行", "duplicate header row"),
    ("错误:", "error:"),
    ("问题", "Problem"),
    ("项", "items"),
    ("项目", "Item"),
    ("（已留少量余量）", " (some headroom kept)"),
    ("（该列在样本中全为空）", "(column empty in the sample)"),
    // ---- 模板匹配（整串锚定；长字面量优先，避免短模板抢先命中）----
    ("  ⚠ 以下字段的原值带时区偏移（如 2024-01-01T08:30:00+08:00），MySQL 的 datetime 不存时区，偏移量已被丢弃、只保留字面时间；如需按时区换算请先在 Excel 里统一：{}", "  ! these columns had timezone offsets (e.g. 2024-01-01T08:30:00+08:00); MySQL datetime has no timezone, so the offset was dropped and only the literal time kept - normalise in Excel first if you need the offset applied: {}"),
    ("  ⚠ 检测到 {} 个百分比格式的单元格：Excel 里显示 12.5%，实际存储值是原始小数 0.125，入库的也是 0.125", "  ! {} cell(s) use percentage formatting: Excel shows 12.5% but the stored value is 0.125, and 0.125 is what goes into the database"),
    ("参数不足：需要 文件名 sheet名 新表名 字段名称所在行 第一个数据所在行 [共几行]，实际给了 {} 个\n详见 --help", "not enough arguments: expected FILE.xlsx SHEET TABLE HEADER_ROW FIRST_DATA_ROW [ROW_COUNT], but got {}\nsee --help"),
    ("  ⚠ 以下字段有 15 位以上的数字，Excel 只能精确保存 15 位，末位可能已被改写为 0，请核对原始数据：{}", "  ! these columns contain numbers with more than 15 digits; Excel keeps only 15 significant digits and the tail may already have become zeros - verify against the source: {}"),
    ("输出目录不存在: {}（请先创建该目录，或用 --out / --err-file / --report 指定别处）", "output directory does not exist: {} (create it first, or point --out / --err-file / --report elsewhere)"),
    ("无法启动 Python 解释器 {}\n  {}\n  可用 --python 指定一个装了 openpyxl 的解释器", "cannot start the Python interpreter {}\n  {}\n  use --python to point at an interpreter with openpyxl installed"),
    ("内容长度 {} 超过 {}({}) 限制（样本未覆盖到，可加 --full-scan 重新扫描或先处理数据）", "length {} exceeds the {}({}) limit (not covered by the sample; try --full-scan or clean the data first)"),
    ("缺少 Python 依赖: {}。请先执行 pip install openpyxl xlrd", "missing Python dependency: {}. Please run: pip install openpyxl xlrd"),
    ("  ⚠ 检测到 {} 个隐藏列（{}{}），隐藏只是不显示，内容仍会被导出；如需排除请先删除该列", "  ! {} hidden column(s) ({}{}); hidden only means not displayed, the content is still exported - delete the column if you want it excluded"),
    ("  {} 数据区合并单元格已补齐 {} 个空单元格（--no-fill-down 可关闭）", "  {} filled {} empty cell(s) from merged ranges in the data area (disable with --no-fill-down)"),
    ("含 {} 个超出 MySQL 日期/时间范围的值，整列按文本保存（样本最长 {} 字符）", "{} values fall outside the MySQL date/time range; the whole column is stored as text (longest sample {} chars)"),
    ("数据区第一行只有这一列有内容（{}），常见于「单位：元」之类的说明；若确为数据请忽略", "only this column has content in the first data row ({}), which often means a note such as \"unit: CNY\"; ignore this if it really is data"),
    ("--add-id {} 与已有字段重名，请换一个列名（如 --add-id id）", "--add-id {} collides with an existing column; pick another name (for example --add-id id)"),
    ("数值 {} 小数位超过 decimal({},{}) 允许的 {} 位（可加大扫描样本后重试）", "number {} has more fraction digits than decimal({},{}) permits: {} (raise the sample size and retry)"),
    ("抽样扫描: 共 {} 行，按 {}% 跳跃抽样，每 {} 行取 1 行，实际判定约 {} 行", "sampled scan: {} rows in total, {}% sampling, 1 row in every {}, about {} rows inspected"),
    ("  ⚠ 以下字段有超出 MySQL 日期/时间范围的值，已整列按文本保存：{}", "  ! these columns contain values outside the MySQL date/time range and were stored as text: {}"),
    ("精度需求 decimal({},{}) 超出 MySQL 上限，改用 double", "decimal({},{}) exceeds the MySQL limit, using double instead"),
    ("第 {} 列及之后有内容（{}），但字段名行只到 {} 列；请补齐字段名或删除多余列", "column {} and beyond have content ({}), but the header row stops at column {}; add header names or delete the extra columns"),
    ("  {} 字段名行位于纵向合并区 {}{}:{}{} 内，已取合并区左上角的内容作为字段名", "  {} header row falls inside the vertically merged range {}{}:{}{}; the top-left value was used as the column name"),
    ("字段名疑似标识类(手机/证件/账号)，按文本保存，样本最长 {} 字符", "the name looks like an identifier (phone / ID / account): stored as text, longest sample {} chars"),
    ("--primary-key {} 在字段名行中不存在。可用字段: {}", "--primary-key {} is not in the header row. Available columns: {}"),
    ("第 {} 行（字段名称所在行）没有任何内容，请检查行号参数是否正确", "row {} (the header row) is empty; check the row-number arguments"),
    ("-- 完成: 共 {} 行, 成功 {} 行, 失败 {} 行, 用时 {}", "-- done: {} rows total, {} succeeded, {} failed, in {}"),
    ("内容超长（最长 {} 字符 / {} 字节），使用 longtext", "very long content ({} chars / {} bytes), using longtext"),
    ("发现 {} 处非标准格式（会造成数据错位，需处理源文件后重跑）：", "{} non-standard spot(s) found (they would misalign the data; fix the source file and run again):"),
    ("无法解析 Excel 结构（Python 助手无有效输出）\n{}", "cannot parse the workbook structure (the Python helper produced no usable output)\n{}"),
    ("  ⚠ 「共几行」={} 超出工作表范围，已自动截断到第 {} 行", "  ! ROW_COUNT={} exceeds the sheet; truncated to row {}"),
    ("xlsxtomysql {} —— Excel → MySQL", "xlsxtomysql {} - Excel to MySQL"),
    ("「第一个数据所在行」({}) 必须大于「字段名称所在行」({})", "FIRST_DATA_ROW ({}) must be greater than HEADER_ROW ({})"),
    ("数值 {} 整数位超过 decimal({},{}) 允许的 {} 位", "number {} has more integer digits than decimal({},{}) permits: {}"),
    ("日期 {} 早于 1900 年，无法写入 datetime", "date year {} is before 1900 and cannot go into datetime"),
    ("读取 Excel 失败（Python 助手退出码 {}）\n{}", "reading Excel failed (Python helper exit code {})\n{}"),
    ("（该列在样本中全为空，另有 {} 个 Excel 错误值）", "(column empty in the sample, plus {} Excel error cells)"),
    ("内容较长（最长 {} 字符 / {} 字节），使用 text", "long content ({} chars / {} bytes), using text"),
    ("包含非数值/非日期内容，降级为文本，样本最长 {} 字符", "contains non-numeric / non-date content, downgraded to text, longest sample {} chars"),
    ("发现 {} 处非标准格式，已中止转换（未生成 sql）。", "{} non-standard spot(s) found; conversion aborted (no SQL written)."),
    ("数据区为空: 起始行 {} 已超出工作表范围（共 {} 行）", "empty data area: start row {} is beyond the sheet ({} rows in total)"),
    ("第 {} 行（字段名称所在行）超出工作表范围（共 {} 行）", "row {} (HEADER_ROW) is beyond the sheet ({} rows in total)"),
    ("\r\x1b[K{} [{}] {}%  {}/{} 行  已用 {}  剩余 {}", "\r\x1b[K{} [{}] {}%  {}/{} rows  elapsed {}  left {}"),
    ("已自动命名为 `{}`（如需自定义请补全 {}{} 单元格）", "renamed to `{}` automatically (fill in cell {}{} to customise)"),
    ("  ⚠ 检测到 {} 个隐藏行，其内容同样会被导出", "  ! {} hidden row(s); their content is exported as well"),
    ("抽样跳跃扫描（每 {} 行取 1 行，判定 {} 行）", "sampled scan (1 row in every {}, {} rows inspected)"),
    ("  {} 字段名行有 {} 处合并单元格，已自动填充{}", "  {} header row has {} merged range(s), filled automatically{}"),
    ("内容: {}；转换时会被当作普通数据行，建议删除", "content: {}; it will be converted as ordinary data, consider deleting it"),
    ("参数过多（{} 个），最后一个是可选的「共几行」", "too many arguments ({}); only the last one (ROW_COUNT) is optional"),
    ("打开 .xls 失败（可能已加密或损坏）: {}", "cannot open .xls (possibly encrypted or corrupt): {}"),
    ("该行内容像分隔符（{}），不是数据；请删除后重试", "this row looks like a divider ({}), not data; delete it and retry"),
    ("含小数的数值，整数位最长 {} 位、小数 {} 位", "decimal values, up to {} integer digits and {} fraction digits"),
    ("无有效数据，默认 varchar(255){}", "no usable data, defaulting to varchar(255){}"),
    ("-- 由 xlsxtomysql {} 生成", "-- generated by xlsxtomysql {}"),
    ("值为 {}，不含日期部分，无法写入 date", "value {} has no date part and cannot go into a date column"),
    ("另有 {} 处提醒（不阻断转换，建议检查）：", "{} additional notice(s) (conversion continues, but please review):"),
    ("整数超出 bigint 范围（样本 {}~{}）", "integers exceed the bigint range (sample {}~{})"),
    ("失败明细（原文件行号，最多显示 {} 条）", "Failed rows (original row numbers, showing up to {})"),
    ("无法启动 Python 解释器 {}\n  {}", "cannot start the Python interpreter {}\n  {}"),
    ("值为 {}，不是整数，无法写入整数型字段", "value {} is not an integer and cannot go into an integer column"),
    ("未知选项 --{}（详见 --help）", "unknown option --{} (see --help)"),
    ("值为 {}，不是数值，无法写入 {} 字段", "value {} is not numeric and cannot go into a {} column"),
    ("值为 {}，含小数，无法写入整数型字段", "value {} has a fraction and cannot go into an integer column"),
    ("数据区末尾有 {} 行空白，已自动忽略", "{} blank row(s) at the end of the data were ignored"),
    ("helper: 未知 mode {}", "helper: unknown mode {}"),
    ("工作表 {} 不存在。可用工作表: {}", "worksheet {} does not exist. Available sheets: {}"),
    ("未知选项 {}（详见 --help）", "unknown option {} (see --help)"),
    ("\r\x1b[K{} 完成: {} {}  用时 {}{}", "\r\x1b[K{} done: {} {}  in {}{}"),
    ("-- 字段数       : {}", "-- Columns      : {}"),
    ("-- 工作表       : {}", "-- Sheet        : {}"),
    ("-- 源文件       : {}", "-- Source file  : {}"),
    ("字段名疑似金额类，样本范围 {}~{}", "the name looks like money, sample range {}~{}"),
    ("-- 生成时间     : {}", "-- Generated at : {}"),
    ("与前面的字段重名(已改名 {})", "duplicate of an earlier column (renamed to {})"),
    ("导出失败行失败（退出码 {}）\n{}", "cannot write the failed-rows workbook (exit code {})\n{}"),
    ("读取阶段异常退出（退出码 {}）", "the reader exited abnormally (exit code {})"),
    ("  ... 其余 {} 行见 {}", "  ... {} more row(s) in {}"),
    ("  检测到 {} 处合并单元格", "  {} merged range(s) detected"),
    ("Python 助手异常: {}", "Python helper failed: {}"),
    ("写入 SQL 尾部失败: {}", "failed to write the SQL trailer: {}"),
    ("打开 .xlsx 失败: {}", "cannot open .xlsx: {}"),
    ("数值 {} 超出 {} 范围({}~{})", "number {} is outside the {} range ({}~{})"),
    ("文本内容，样本最长 {} 字符", "text content, longest sample {} chars"),
    ("-- 字段名所在行 : {}", "-- Header row   : {}"),
    ("-- 第一条数据行 : {}", "-- First data   : {}"),
    ("{} {}  （{} 行 × {} 列）", "{} {}  ({} rows x {} cols)"),
    ("{} 完成: {} 行  用时 {}{}", "{} done: {} rows  in {}{}"),
    ("值为 '{}'，不是合法数值", "value '{}' is not a valid number"),
    ("值为 '{}'，无法解析为 {}", "value '{}' cannot be parsed as {}"),
    ("全部为整数，样本范围 {}~{}", "all integers, sample range {}~{}"),
    ("打开 .xls 失败: {}", "cannot open .xls: {}"),
    ("数值 '{}' 不是有限数值", "number '{}' is not finite"),
    ("（其中 {} 处为纵向合并）", " ({} of them vertical)"),
    ("写入 SQL 失败: {}", "failed to write SQL: {}"),
    ("报表写入失败: {} ({})", "cannot write the report {}: {}"),
    ("第 {} 行 / 第 {} 列", "row {} / col {}"),
    ("--{} 后面需要一个值", "--{} requires a value"),
    ("值为 {}，无法解析为 {}", "value {} cannot be parsed as {}"),
    ("  可用工作表: {}", "  available sheets: {}"),
    ("导出失败行失败: {}", "cannot write the failed-rows workbook: {}"),
    ("生成 {} 失败: {}", "failed to create {}: {}"),
    ("--{} 需要整数", "--{} expects an integer"),
    ("{} {}%  {}/{} 行", "{} {}%  {}/{} rows"),
    ("全量扫描 {} 行", "full scan of {} rows"),
    ("报表已保存: {}", "report saved: {}"),
    ("文件不存在: {}", "file not found: {}"),
    ("无法写入 {}: {}", "cannot write {}: {}"),
    ("{}（{} 个）", "{} ({} of them)"),
    ("完成: {}", "done: {}"),
    ("第 {} 行", "row {}"),
    ("错误: {}", "error: {}"),
    ("{} {} 行", "{} {} rows"),
];
// ==== I18N CATALOG END ====

/// 精确查表：中文原文 → 英文。查不到原样返回。
fn en_of(zh: &str) -> &str {
    for (z, e) in CATALOG {
        if *z == zh {
            return e;
        }
    }
    zh
}

/// 文案翻译（无参数）。中文环境或串里没有中文时原样返回。
fn t(s: &str) -> String {
    if lang_zh() || !needs_translation(s) { s.to_string() } else { en_of(s).to_string() }
}

/// 文案翻译（带参数）：先取英文模板，再把模板里的 {} 依次填上。
fn tf(tmpl: &str, args: &[&dyn std::fmt::Display]) -> String {
    let text: String = if lang_zh() || !needs_translation(tmpl) {
        tmpl.to_string()
    } else {
        en_of(tmpl).to_string()
    };
    fill_args(&text, args)
}

/// 把文本里的 {} 按顺序替换成实参。{:.0} 这类带格式说明的不算占位符。
fn fill_args(text: &str, args: &[&dyn std::fmt::Display]) -> String {
    let mut out = String::with_capacity(text.len() + 16 * args.len());
    let mut it = text.chars().peekable();
    let mut ai = 0usize;
    while let Some(c) = it.next() {
        if c == '{' && it.peek() == Some(&'}') {
            it.next();
            if let Some(a) = args.get(ai) {
                out.push_str(&a.to_string());
            }
            ai += 1;
        } else {
            out.push(c);
        }
    }
    out
}

/// 模板匹配：pat 里 {} 是通配段。整串锚定（首尾字面量必须对齐），
/// 所以"第 12 行"绝不会被"第 {} 行 / 第 {} 列"命中。
fn match_template<'a>(pat: &str, s: &'a str) -> Option<Vec<&'a str>> {
    let segs: Vec<&str> = pat.split("{}").collect();
    if segs.len() < 2 {
        return None;
    }
    let first = segs[0];
    let last = segs[segs.len() - 1];
    if !s.starts_with(first) || !s.ends_with(last) {
        return None;
    }
    let end = match s.len().checked_sub(last.len()) {
        Some(e) if e >= first.len() => e,
        _ => return None,
    };
    let mut caps: Vec<&str> = Vec::with_capacity(segs.len() - 1);
    let mut pos = first.len();
    for (k, seg) in segs.iter().enumerate().skip(1) {
        if k == segs.len() - 1 {
            caps.push(&s[pos..end]);
            break;
        }
        let idx = s[pos..end].find(seg)?;
        caps.push(&s[pos..pos + idx]);
        pos = pos + idx + seg.len();
    }
    Some(caps)
}

/// 输出边界翻译：先精确、再模板。有多个模板都命中时取字面量最长的那个
/// （"{} {} 行"和"{} 完成: {} 行  用时 {}{}"相比，后者更具体，应该优先）。
fn tr(s: &str) -> String {
    if lang_zh() || s.is_empty() || !needs_translation(s) {
        return s.to_string();
    }
    tr_depth(s, 0)
}

fn tr_depth(s: &str, depth: usize) -> String {
    if s.is_empty() || !needs_translation(s) {
        return s.to_string();
    }
    if let Some((_, en)) = CATALOG.iter().find(|(z, _)| *z == s) {
        return en.to_string();
    }
    if depth >= 4 {
        return s.to_string();
    }
    // 找字面量最长（最具体）的命中模板
    let mut best: Option<(&str, Vec<&str>, usize)> = None;
    for (zh, en) in CATALOG {
        if !zh.contains("{}") {
            continue;
        }
        if let Some(caps) = match_template(zh, s) {
            let lit: usize = zh.split("{}").map(|x| x.chars().count()).sum();
            if best.as_ref().map(|b| lit > b.2).unwrap_or(true) {
                best = Some((en, caps, lit));
            }
        }
    }
    match best {
        None => s.to_string(),
        Some((en, caps, _)) => {
            // 实参本身可能还是中文文案（比如"无有效数据，默认 varchar(255)（该列…）"），
            // 所以逐段再翻一层。
            let segs: Vec<&str> = en.split("{}").collect();
            let mut out = String::with_capacity(en.len() + 24 * caps.len());
            for (i, seg) in segs.iter().enumerate() {
                out.push_str(seg);
                if i < caps.len() {
                    out.push_str(&tr_depth(caps[i], depth + 1));
                }
            }
            out
        }
    }
}

/// 列表连接符：中文顿号 / 英文逗号
fn sep() -> &'static str {
    if lang_zh() { "、" } else { ", " }
}

/// 子句分隔符：中文分号 / 英文分号加空格
fn sep2() -> &'static str {
    if lang_zh() { "；" } else { "; " }
}

fn help_text() -> &'static str {
    if lang_zh() { HELP_ZH } else { HELP_EN }
}

fn man_text() -> &'static str {
    if lang_zh() { MAN_ZH } else { MAN_EN }
}

/// 恢复 Unix 默认的 SIGPIPE 行为。
/// Rust 运行时默认忽略 SIGPIPE，于是 `xlsxtomysql --man | head` 这类管道会被下游关闭，
/// 后续写 stdout 直接 panic 并打印一堆 backtrace。改回默认后进程安静退出，和普通
/// Unix 命令一致。
#[cfg(unix)]
fn restore_sigpipe() {
    extern "C" {
        fn signal(sig: i32, handler: usize) -> usize;
    }
    const SIGPIPE: i32 = 13;
    const SIG_DFL: usize = 0;
    unsafe {
        signal(SIGPIPE, SIG_DFL);
    }
}

#[cfg(not(unix))]
fn restore_sigpipe() {}

/// Windows 控制台：切到 UTF-8 代码页并打开 ANSI 转义（彩色输出要用）
#[cfg(windows)]
fn console_setup() {
    #[link(name = "kernel32")]
    extern "system" {
        fn SetConsoleOutputCP(code_page: u32) -> i32;
        fn SetConsoleCP(code_page: u32) -> i32;
        fn GetStdHandle(which: u32) -> *mut core::ffi::c_void;
        fn GetConsoleMode(handle: *mut core::ffi::c_void, mode: *mut u32) -> i32;
        fn SetConsoleMode(handle: *mut core::ffi::c_void, mode: u32) -> i32;
    }
    const STD_OUTPUT_HANDLE: u32 = -11i32 as u32;
    const ENABLE_VIRTUAL_TERMINAL_PROCESSING: u32 = 0x0004;
    unsafe {
        SetConsoleOutputCP(65001);
        SetConsoleCP(65001);
        let h = GetStdHandle(STD_OUTPUT_HANDLE);
        let mut mode: u32 = 0;
        if !h.is_null() && GetConsoleMode(h, &mut mode) != 0 {
            SetConsoleMode(h, mode | ENABLE_VIRTUAL_TERMINAL_PROCESSING);
        }
    }
}

#[cfg(not(windows))]
fn console_setup() {}

const HELP_EN: &str = r#"xlsxtomysql - convert Excel into MySQL CREATE TABLE + INSERT statements (single-file Rust build)

Usage:
  xlsxtomysql FILE.xlsx  SHEET  TABLE  HEADER_ROW  FIRST_DATA_ROW  [ROW_COUNT]

  FILE.xlsx       Excel 2007+ (.xlsx) or Excel 2003 (.xls)
  SHEET           worksheet name
  TABLE           new table name; also the output file name <TABLE>.sql
  HEADER_ROW      1-based row number holding the column names
  FIRST_DATA_ROW  1-based row number of the first data row
  ROW_COUNT       optional, default: until the end of the file

Main options:
  --out PATH              output .sql path (default: next to the source file as <TABLE>.sql)
  --err-file PATH         failed-rows workbook (default: next to the source file as errrows.xlsx)
  --report PATH           also save the console report as a text file
  --force                 keep converting even when the sheet format is non-standard
  --scan-only             scan and print the inferred types, do not write SQL
  --no-hints              disable column-name based hints (phone numbers, money, ...)
  --hints                 enable them again (on by default; useful to override an earlier --no-hints)
  --full-scan             always scan every row (slow on large files)
  --sample-ratio F        sampling ratio, default 0.10
  --sample-threshold N    row count above which sampled skip scanning kicks in, default 2000
  --sample-min N          minimum sampled rows, default 1000
  --sample-max N          maximum sampled rows, default 20000
  --seed N                random seed for the sampling start (reproducible runs)
  --batch-size N          rows per INSERT statement, default 200
  --empty-as-null M       auto | always | never, default auto
  --varchar-max N         varchar ceiling, default 1000
  --text-max N            text ceiling, above which longtext is used, default 16000
  --max-ident N           maximum column-name length, default 64
  --merge-scan M          auto | on | off, default auto
  --merge-scan-limit N    max file size in MB for merge scanning, default 64
  --no-fill-down          do not fill merged cells in the data area
  --add-id NAME           append an auto-increment primary key column
  --primary-key COL       use COL as the primary key
  --drop-table            emit DROP TABLE IF EXISTS
  --insert-ignore         emit INSERT IGNORE
  --table-comment S       table comment
  --not-null              add NOT NULL to every column
  --print-errors N        max failed rows shown on the console, default 20
  --progress off          disable the progress bar (--no-progress is equivalent)
  --no-color              disable colored output (auto-off when piping)
  --python PATH           Python interpreter used to read Excel
  --dump-python           print the embedded Python helper source and exit
  --man                   show the full manual (option details, type rules, exit codes)
  --lang zh|en            force Chinese/English output (default: auto-detect)
  -h, --help              show this help
  -V, --version           show version

Language:
  Chinese environments (zh) get Chinese output, everything else gets English.
  Order: --lang > XLSXTOMYSQL_LANG > LC_ALL > LC_MESSAGES > LANG;
  when none is set: English on Unix, the system UI language on Windows.
  Comments inside the generated .sql follow the same language.
"#;

const MAN_ZH: &str = r#"xlsxtomysql 手册
======================================================================

名称
    xlsxtomysql —— 把 Excel（.xlsx / .xls）转换成 MySQL 建表语句与插入语句。

用法
    xlsxtomysql 文件名.xlsx sheet名 新表名 字段名称所在行 第一个数据所在行 [共几行] [选项...]

    -h / --help   精简用法（选项速查）
    --man         本手册（参数细节、类型推断规则、退出码、已知限制）
    -V           版本号

位置参数
    文件名.xlsx       .xlsx（Excel 2007+）或 .xls（Excel 2003）。
                      读 .xls 需要 Python 侧装有 xlrd。
    sheet名           工作表名，必须完全一致（错误提示里会列出可选工作表）。
    新表名            既是 CREATE TABLE 的表名，也是默认输出文件名 <新表名>.sql。
    字段名称所在行     1 起算。允许落在纵向合并区内（程序回读合并区左上角的值）。
    第一个数据所在行   1 起算，必须大于字段名称所在行。
    共几行            可选。只转换这么多行数据；超出工作表范围会自动截断并提示。

选项
    输出与行为
      --out PATH            输出 sql 的路径（默认：源文件同目录/<新表名>.sql）
      --err-file PATH       失败行文件（默认：源文件同目录/errrows.xlsx）
      --report PATH         把控制台报表另存一份纯文本（便于存档/发邮件）
      --scan-only           只预扫描并输出「结果表 2」，不生成 sql
      --force               发现非标准格式时不再阻断，强行转换
    INSERT 语句
      --batch-size N        每条 INSERT 的行数，默认 200
      --insert-ignore       生成 INSERT IGNORE
      --empty-as-null M     auto | always | never，空单元格是否写 NULL，默认 auto
      --not-null            所有字段加 NOT NULL
      --drop-table          生成 DROP TABLE IF EXISTS
      --add-id NAME         追加一个 bigint 自增主键列（并作为主键）
      --primary-key COL     指定主键列（列名取自字段名行）
      --table-comment S     表注释（写入 COMMENT='...'）
    类型推断
      --varchar-max N       varchar 上限，默认 1000；超过则考虑 text/longtext
      --text-max N          text 上限，默认 16000；超过改用 longtext
      --max-ident N         字段名最大长度，默认 64（超长会截断并说明）
      --no-hints            关闭字段名语义推断（手机号/证件号/金额等按名字给类型）
      --hints               重新打开（默认就是开的，主要用来覆盖前面的 --no-hints）
    扫描与抽样
      --full-scan           强制全量扫描（最准，大文件慢）
      --sample-threshold N  超过多少行启用抽样跳跃扫描，默认 2000
      --sample-ratio F      抽样比例，默认 0.10（10%）
      --sample-min N        抽样行数下限，默认 1000
      --sample-max N        抽样行数上限，默认 20000
      --seed N              抽样起点随机种子，便于复现同一次扫描
    合并单元格
      --merge-scan M        auto | on | off，是否扫描合并区域，默认 auto
      --merge-scan-limit N  合并扫描的大小上限（MB），默认 64
      --no-fill-down        不自动补齐数据区合并单元格
    其它
      --print-errors N      控制台最多显示多少条失败行，默认 20
      --progress off        关闭进度条；--no-progress 等价
      --no-color            关闭彩色输出（管道输出会自动关闭）
      --python PATH         指定读取 Excel 的 Python 解释器
      --dump-python         打印内嵌的 Python 助手源码后退出
      --man                 本手册
      --lang zh|en          强制输出语言
      -h, --help / -V, --version

退出码
    0   成功（即使有失败行，只要生成了 sql 就算成功，失败行会列在报表里）
    1   运行错误（读文件、写文件、Python 助手异常等）
    2   非标准格式阻断：源文件需要先处理，或加 --force 强行继续
    3   参数错误或引用错误（文件不存在、工作表名不对、行号越界等）

类型推断
    先按整列样本判定种类，再给 MySQL 类型：
      全整数        → tinyint / smallint / mediumint / int / bigint（按数值范围挑最小够用的）
      含小数        → decimal(p,s)，精度超出 MySQL 上限（65/30）时改用 double
      科学计数法     → double
      真假值        → tinyint(1)
      纯日期        → date      纯时间 → time      纯日期时间 → datetime
      日期+日期时间  → datetime（统一）
      日期与时间混杂 → varchar（按文本保存，避免失真）
      固定文本      → varchar(n)，n 按样本最长值向上取「整齐档位」
      长文本        → text；超过 65535 字节或超过 --text-max 则 longtext
      全空          → varchar(255)，并在「推断依据」里说明
    字段名提示（--no-hints 可关）：
      名字含手机/电话/证件/账号/学号/邮箱/邮编等 → 即使整列是数字也按文本存
      名字含金额/价格/费用/工资等 → 倾向 decimal(p,2)
    注意：text / longtext 的判断按 UTF-8 字节数（中文一个字 3 字节）。

抽样扫描
    行数超过 --sample-threshold（默认 2000）时，先无条件判断开头一段，然后按「质数步长」
    跳跃抽样到大约 --sample-ratio（默认 10%）的行。用质数步长是为了避免与数据自身周期
    共振（否则固定周期数据会系统性漏掉一部分取值，把类型判错）。抽样行数在
    --sample-min / --sample-max 之间兜底。扫描阶段识别出的小数位/长度不够用时，
    转换阶段会逐行报错并计入失败行；可加 --full-scan 或调大 --sample-ratio 重跑。
    想复现同一次抽样，用 --seed N。

特殊表格
    程序会检查并给出「位置 + 问题 + 说明」，需要处理完再重跑（除非加 --force）：
      * 数据区中间整行为空（会打断数据）→ 阻断
      * 分隔线/装饰行（如 ----、====）→ 阻断
      * 重复的字段名行（表头出现两次）→ 阻断
      * 数据超出字段名范围（字段名行只到 A-C，数据却到 E）→ 阻断
      * 数据区为空 / 字段名行没有内容 → 阻断
      * 疑似说明行（「单位：元」之类）、疑似合计/小计行、末尾空行 → 只提醒
    会自动处理并提示：
      * 数据区合并单元格：按左上角内容补齐整块（--no-fill-down 关闭）
      * 字段名行位于纵向合并区：取合并区左上角那一行的值
      * 隐藏行列：内容仍会被导出（只提示，不改数据）
      * 单元格内图片（WPS DISPIMG / Excel 单元格图片）：无法写入 SQL，对应单元格变空
      * 百分比格式：Excel 显示 12.5%，实际存的是 0.125，入库也是 0.125
      * 带时区偏移的日期时间（…+08:00）：MySQL 不存时区，偏移被丢弃，只保留字面时间
      * 15 位以上的数字：Excel 只精确保存 15 位，末位可能已被改写为 0
      * 超出 MySQL 日期/时间范围的值：整列按文本保存
    不会猜测的写法（一律按文本保存，需自行处理）：12/31/2024 这类有歧义的日期、
    €1,234.56 这类货币表达、文本形式的 12.5%、全角数字、带前导零的编号。

报表
    控制台输出两张表：
      结果表 1：行数、成功/失败/警告/跳过、扫描方式、耗时、输出文件、失败行文件；
                有失败行时附「失败明细」与「失败原因归类」。
      结果表 2：每个新字段名、MySQL 类型、推断依据、字段名处理（改名/截断/非空说明）。
    失败行（含原行号、原值、原因）导出到 errrows.xlsx，便于逐行修正。
    --report PATH 可把控制台报表另存一份文本。

环境变量
    XLSXTOMYSQL_PYTHON   读取 Excel 用的 Python 解释器（等价 --python）
    XLSXTOMYSQL_LANG     输出语言：zh / en（优先级高于系统 locale）
    LC_ALL / LC_MESSAGES / LANG   语言自动检测（zh 开头 → 中文，其它 → 英文）

示例
    # 标准表：第 1 行字段名，第 2 行起是数据
    xlsxtomysql 学生信息.xlsx 学生信息 students 1 2

    # 只要前 100 行，带上 drop table 与主键
    xlsxtomysql 名单.xlsx Sheet1 t_list 2 3 100 --drop-table --add-id id

    # 大文件先看类型推断结果
    xlsxtomysql 订单.xlsx 明细 t_order 1 2 --scan-only

    # 源文件有多余列、又必须转：强行继续
    xlsxtomysql 报表.xlsx Sheet1 t_rpt 1 2 --force

已知限制
    * 图片、批注、超链接、图表不是单元格值，不参与导出；浮动图片若遮挡数据行，
      程序无法识别（浮动对象没有行列归属），这类表请人工核对。
    * 只读取公式的当前缓存结果；整列为空时请先用 Excel 打开保存一次。
    * 抽样扫描可能漏掉少量极端值，转换阶段会如实报错并计入失败行。转换前请保留
      源文件副本，生成的 sql 建议先导入测试库确认。
    * .xls 依赖 Python 的 xlrd；.xlsx 依赖 openpyxl。缺依赖时按提示安装。
    * 加密/损坏的工作簿无法读取，需要先解密或修复。

许可
    MIT License。
"#;

const MAN_EN: &str = r#"xlsxtomysql manual
======================================================================

Name
    xlsxtomysql - convert Excel (.xlsx / .xls) into MySQL CREATE TABLE + INSERT statements.

Usage
    xlsxtomysql FILE.xlsx SHEET TABLE HEADER_ROW FIRST_DATA_ROW [ROW_COUNT] [options...]

    -h / --help   short usage (option cheat sheet)
    --man         this manual (option details, type rules, exit codes, limitations)
    -V            version

Positional arguments
    FILE.xlsx      .xlsx (Excel 2007+) or .xls (Excel 2003). Reading .xls needs xlrd
                   on the Python side.
    SHEET          worksheet name; must match exactly (available names are listed on error).
    TABLE          table name for CREATE TABLE, also the default output file <TABLE>.sql.
    HEADER_ROW     1-based. May sit inside a vertical merge (the top-left value is read back).
    FIRST_DATA_ROW 1-based, must be greater than HEADER_ROW.
    ROW_COUNT      optional; convert only this many data rows. Values beyond the sheet are
                   truncated with a notice.

Options
    Output and behaviour
      --out PATH             output .sql path (default: next to the source, <TABLE>.sql)
      --err-file PATH        failed-rows workbook (default: next to the source, errrows.xlsx)
      --report PATH          also save the console report as plain text
      --scan-only            pre-scan and print result table 2 only, no SQL
      --force                do not stop on non-standard sheet formats
    INSERT statements
      --batch-size N         rows per INSERT, default 200
      --insert-ignore        emit INSERT IGNORE
      --empty-as-null M      auto | always | never, whether empty cells become NULL, default auto
      --not-null             add NOT NULL to every column
      --drop-table           emit DROP TABLE IF EXISTS
      --add-id NAME          append a bigint auto-increment primary key column
      --primary-key COL      use COL as the primary key (name from the header row)
      --table-comment S      table comment (COMMENT='...')
    Type inference
      --varchar-max N        varchar ceiling, default 1000
      --text-max N           text ceiling above which longtext is used, default 16000
      --max-ident N          max column-name length, default 64 (longer names are truncated)
      --no-hints             disable column-name hints (phone / ID / money / ...)
      --hints                enable them again (on by default; useful to override an earlier --no-hints)
    Scanning and sampling
      --full-scan            scan every row (most accurate, slow on big files)
      --sample-threshold N   row count above which sampled skip scanning starts, default 2000
      --sample-ratio F       sampling ratio, default 0.10
      --sample-min N         minimum sampled rows, default 1000
      --sample-max N         maximum sampled rows, default 20000
      --seed N               seed for the sampling start (reproducible scans)
    Merged cells
      --merge-scan M         auto | on | off, default auto
      --merge-scan-limit N   size limit in MB for merge scanning, default 64
      --no-fill-down         do not fill merged cells in the data area
    Misc
      --print-errors N       max failed rows shown on the console, default 20
      --progress off         disable the progress bar (--no-progress is equivalent)
      --no-color             disable colored output (auto-off when piping)
      --python PATH          Python interpreter used to read Excel
      --dump-python          print the embedded Python helper source and exit
      --man                  this manual
      --lang zh|en           force the output language
      -h, --help / -V, --version

Exit codes
    0   success (failed rows do not change this, as long as SQL was generated; they are listed)
    1   runtime error (reading/writing files, Python helper failure, ...)
    2   blocked: non-standard sheet format; fix the source or pass --force
    3   bad arguments or bad references (missing file, wrong sheet name, row out of range)

Type inference
    The whole column sample is classified first, then mapped to a MySQL type:
      integers only     -> tinyint / smallint / mediumint / int / bigint (smallest that fits)
      with decimals     -> decimal(p,s); beyond MySQL limits (65/30) -> double
      scientific        -> double
      boolean           -> tinyint(1)
      dates only        -> date      times only -> time      datetimes only -> datetime
      dates + datetimes -> datetime (unified)
      dates mixed with times -> varchar (kept as text to avoid distortion)
      text              -> varchar(n) rounded up to a tidy ladder step
      long text         -> text; over 65535 bytes or over --text-max -> longtext
      empty column      -> varchar(255), explained in the "inference basis" column
    Column-name hints (disable with --no-hints):
      phone / ID / account / student-no / e-mail / postcode -> stored as text even if numeric
      amount / price / fee / salary -> prefers decimal(p,2)
    Note: text vs longtext is judged in UTF-8 bytes (a CJK character is 3 bytes).

Sampled skip scanning
    Above --sample-threshold (default 2000) rows the tool always judges a warm-up block first,
    then skips through the rest with a prime step until roughly --sample-ratio (default 10%)
    of the rows are judged. A prime step avoids resonance with periodic data, which would
    otherwise systematically miss values and misjudge types. The sampled count is bounded by
    --sample-min / --sample-max. Values the sample missed fail during conversion and are
    reported per row; rerun with --full-scan or a larger --sample-ratio. Use --seed N to
    reproduce a specific sampling.

Special sheet layouts
    Detected, reported with position + problem + explanation, and blocking unless --force:
      * a fully blank row in the middle of the data (breaks the dataset)
      * divider/decoration rows (----, ====)
      * a duplicated header row
      * data beyond the header range (headers end at C but data reaches E)
      * empty data area / empty header row
    Reported but not blocking: suspected note rows ("Unit: CNY"), total/subtotal rows,
    trailing blank rows.
    Handled automatically with a notice:
      * merged cells in the data area are filled from the top-left value (--no-fill-down off)
      * a header row inside a vertical merge takes the top-left value of the merge
      * hidden rows/columns are still exported (noticed, data untouched)
      * in-cell images (WPS DISPIMG / Excel cell pictures) cannot be written to SQL; the cell
        becomes empty
      * percentage formatting: Excel shows 12.5% but stores 0.125, which is what gets inserted
      * datetimes with a timezone offset (...+08:00): MySQL stores no timezone, the offset is
        dropped and the literal time kept
      * numbers with more than 15 digits: Excel only keeps 15 exactly; the tail may be zeroed
      * values outside the MySQL date/time range: the whole column is stored as text
    Never guessed (kept as text, handle them yourself): ambiguous dates like 12/31/2024,
    currency strings like €1,234.56, textual 12.5%, full-width digits, zero-padded codes.

Reports
    Two tables are printed:
      Result table 1: rows, ok/failed/warned/skipped, scan mode, elapsed, output file,
                      failed-rows file; plus failure details and failure reason groups.
      Result table 2: each new column name, MySQL type, inference basis, header handling.
    Failed rows (source row number, original values, reason) go to errrows.xlsx so they can be
    fixed one by one. --report PATH also saves the console report as text.

Environment variables
    XLSXTOMYSQL_PYTHON   Python interpreter used to read Excel (same as --python)
    XLSXTOMYSQL_LANG     output language: zh / en (wins over the system locale)
    LC_ALL / LC_MESSAGES / LANG   automatic detection (zh* -> Chinese, otherwise English)

Examples
    # standard sheet: header on row 1, data from row 2
    xlsxtomysql students.xlsx Sheet1 students 1 2

    # only the first 100 rows, with DROP TABLE and a surrogate key
    xlsxtomysql roster.xlsx Sheet1 t_list 2 3 100 --drop-table --add-id id

    # inspect the inferred types of a large file first
    xlsxtomysql orders.xlsx detail t_order 1 2 --scan-only

    # extra columns in the source, but conversion must go on
    xlsxtomysql report.xlsx Sheet1 t_rpt 1 2 --force

Limitations
    * Pictures, comments, hyperlinks and charts are not cell values and are not exported.
      Floating pictures covering data rows cannot be detected (floating objects have no cell
      anchor) - check such sheets by hand.
    * Only cached formula results are read; open and re-save the file in Excel when a formula
      column comes out empty.
    * Sampling can miss extreme values; those rows then fail individually and are reported.
      Keep a copy of the source file and import the generated SQL into a test database first.
    * .xls needs xlrd on the Python side; .xlsx needs openpyxl.
    * Encrypted or corrupt workbooks cannot be read; decrypt or repair them first.

License
    MIT License.
"#;

// ===========================================================================
// 配色
// ===========================================================================
#[derive(Clone, Copy)]
struct Style {
    color: bool,
}

impl Style {
    fn wrap(&self, code: &str, s: &str) -> String {
        // 带色输出的文案也走翻译：颜色码是额外套上去的，先翻再套
        let text = tr(s);
        if self.color { format!("\x1b[{}m{}\x1b[0m", code, text) } else { text }
    }
    fn bold(&self, s: &str) -> String { self.wrap("1", s) }
    fn dim(&self, s: &str) -> String { self.wrap("2", s) }
    fn red(&self, s: &str) -> String { self.wrap("31", s) }
    fn yellow(&self, s: &str) -> String { self.wrap("33", s) }
    fn cyan(&self, s: &str) -> String { self.wrap("36", s) }
}

// ===========================================================================
// 输出（同时攒一份报表文本）
// ===========================================================================
struct Out {
    buf: Option<String>,
}

impl Out {
    fn new(report: bool) -> Self {
        Out { buf: if report { Some(String::new()) } else { None } }
    }
    fn line(&mut self, s: &str) {
        // 所有报表行都在这里出门，翻译放这里最不容易漏
        let text = tr(s);
        println!("{}", text);
        if let Some(b) = self.buf.as_mut() {
            b.push_str(&text);
            b.push('\n');
        }
    }
    fn blank(&mut self) {
        self.line("");
    }
    fn save(&self, path: &str) -> io::Result<()> {
        if let Some(b) = self.buf.as_ref() {
            std::fs::write(path, b)?;
        }
        Ok(())
    }
}

// ===========================================================================
// 进度条（写 stderr，不污染可重定向的报表 stdout）
// ===========================================================================
struct Bar {
    label: String,
    total: i64,
    n: i64,
    enabled: bool,
    tty: bool,
    t0: Instant,
    last: f64,
    last_decile: i64,
    done: bool,
}

impl Bar {
    fn new(label: &str, enabled: bool) -> Self {
        Bar {
            label: label.to_string(),
            total: 0,
            n: 0,
            enabled,
            tty: io::stderr().is_terminal(),
            t0: Instant::now(),
            last: 0.0,
            last_decile: -1,
            done: false,
        }
    }
    fn set_total(&mut self, t: i64) {
        if t > 0 {
            self.total = t;
        }
    }
    fn set(&mut self, n: i64) {
        self.n = n;
        self.render(false);
    }
    fn render(&mut self, force: bool) {
        if !self.enabled || self.done {
            return;
        }
        let now = self.t0.elapsed().as_secs_f64();
        if !self.tty {
            let decile = if self.total > 0 { self.n * 10 / self.total } else { 0 };
            if force || decile > self.last_decile {
                self.last_decile = decile;
                if self.total > 0 {
                    let pct = format!("{:5.1}", self.n as f64 / self.total as f64 * 100.0);
                    eprintln!("{}", tf("{} {}%  {}/{} 行",
                              &[&self.label, &pct, &human_num(self.n), &human_num(self.total)]));
                } else {
                    eprintln!("{}", tf("{} {} 行", &[&self.label, &human_num(self.n)]));
                }
            }
            return;
        }
        if !force && now - self.last < 0.08 {
            return;
        }
        self.last = now;
        let width = 26usize;
        let ratio = if self.total > 0 { (self.n as f64 / self.total as f64).min(1.0) } else { 0.0 };
        let filled = (width as f64 * ratio).round() as usize;
        let bar = format!("{}{}", "█".repeat(filled), "░".repeat(width - filled));
        let speed = if now > 0.0 { self.n as f64 / now } else { 0.0 };
        let eta = if speed > 0.0 && self.total > 0 { (self.total - self.n) as f64 / speed } else { 0.0 };
        let pct = format!("{:5.1}", ratio * 100.0);
        eprint!("{}", tf("\r\x1b[K{} [{}] {}%  {}/{} 行  已用 {}  剩余 {}",
            &[&self.label, &bar, &pct, &human_num(self.n), &human_num(self.total),
              &fmt_dur(now), &fmt_dur(eta.max(0.0))]));
        let _ = io::stderr().flush();
    }
    fn finish(&mut self, note: &str) {
        if !self.enabled || self.done {
            return;
        }
        self.render(true);
        self.done = true;
        let elapsed = self.t0.elapsed().as_secs_f64();
        let tail = if note.is_empty() { String::new() } else { format!("  {}", note) };
        let unit = if self.total > 0 { t("行") } else { t("项") };
        let el = fmt_dur(elapsed);
        if self.tty {
            eprintln!("{}", tf("\r\x1b[K{} 完成: {} {}  用时 {}{}",
                      &[&self.label, &human_num(self.n), &unit, &el, &tail]));
        } else {
            eprintln!("{}", tf("{} 完成: {} 行  用时 {}{}",
                      &[&self.label, &human_num(self.n), &el, &tail]));
        }
        let _ = io::stderr().flush();
    }
}


// ===========================================================================
// 单元格
// ===========================================================================
#[derive(Clone, Copy, PartialEq, Eq, Hash, Debug)]
enum Kind { Empty, Int, Dec, Sci, Bool, Text, Date, Time, DateTime, Err }

#[derive(Clone, Debug)]
enum Cell {
    Empty,
    Int(i128),
    Dec(f64),
    Sci(f64),
    Bool(bool),
    Text(String),
    Date(String),
    Time(String),
    DateTime(String),
    Err(String),
    /// 日期/时间超出 MySQL 可表示范围，按文本保存（类型推断时也当文本）
    Oor(String),
}

impl Cell {
    fn kind(&self) -> Kind {
        match self {
            Cell::Empty => Kind::Empty,
            Cell::Int(_) => Kind::Int,
            Cell::Dec(_) => Kind::Dec,
            Cell::Sci(_) => Kind::Sci,
            Cell::Bool(_) => Kind::Bool,
            Cell::Text(_) | Cell::Oor(_) => Kind::Text,
            Cell::Date(_) => Kind::Date,
            Cell::Time(_) => Kind::Time,
            Cell::DateTime(_) => Kind::DateTime,
            Cell::Err(_) => Kind::Err,
        }
    }
    fn is_empty(&self) -> bool {
        matches!(self, Cell::Empty)
    }
    fn text(&self) -> String {
        match self {
            Cell::Empty => String::new(),
            Cell::Int(v) => v.to_string(),
            Cell::Dec(f) | Cell::Sci(f) => fmt_float(*f),
            Cell::Bool(b) => if *b { "1".into() } else { "0".into() },
            Cell::Text(s) | Cell::Date(s) | Cell::Time(s) | Cell::DateTime(s)
            | Cell::Err(s) | Cell::Oor(s) => s.clone(),
        }
    }
}

const EMPTY_CELL: Cell = Cell::Empty;

fn fmt_float(f: f64) -> String {
    if f.fract() == 0.0 && f.abs() < 1e16 {
        format!("{:.1}", f)
    } else {
        format!("{}", f)
    }
}

fn unescape(s: &str) -> String {
    if !s.contains('\\') {
        return s.to_string();
    }
    let mut out = String::with_capacity(s.len());
    let mut it = s.chars();
    while let Some(c) = it.next() {
        if c == '\\' {
            match it.next() {
                Some('t') => out.push('\t'),
                Some('n') => out.push('\n'),
                Some('r') => out.push('\r'),
                Some('\\') => out.push('\\'),
                Some(other) => out.push(other),
                None => out.push('\\'),
            }
        } else {
            out.push(c);
        }
    }
    out
}

fn escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '\\' => out.push_str("\\\\"),
            '\t' => out.push_str("\\t"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            _ => out.push(c),
        }
    }
    out
}

fn parse_cell(tok: &str) -> Cell {
    let (k, rest) = match tok.split_once(':') {
        Some(x) => x,
        None => (tok, ""),
    };
    let v = unescape(rest);
    match k {
        "e" => EMPTY_CELL,
        "i" => match v.parse::<i128>() {
            Ok(n) => Cell::Int(n),
            Err(_) => Cell::Text(v),
        },
        "d" => match v.parse::<f64>() {
            Ok(f) if f.is_finite() => Cell::Dec(f),
            _ => Cell::Text(v),
        },
        "f" => match v.parse::<f64>() {
            Ok(f) if f.is_finite() => Cell::Sci(f),
            _ => Cell::Text(v),
        },
        "b" => Cell::Bool(v == "1"),
        "s" => Cell::Text(v),
        "D" => Cell::Date(v),
        "T" => Cell::Time(v),
        "M" => Cell::DateTime(v),
        "x" => Cell::Err(v),
        "o" => Cell::Oor(v),
        _ => Cell::Text(v),
    }
}

// ===========================================================================
// 参数
// ===========================================================================
struct Opts {
    path: String,
    sheet: String,
    table: String,
    header_row: i64,
    data_row: i64,
    rows: Option<i64>,

    out: Option<String>,
    err_file: Option<String>,
    report: Option<String>,

    force: bool,
    hints: bool,
    progress: bool,
    color: bool,
    scan_only: bool,

    batch_size: i64,
    sample_ratio: f64,
    sample_threshold: i64,
    sample_min: i64,
    sample_max: i64,
    full_scan: bool,
    seed: Option<u64>,

    add_id: Option<String>,
    primary_key: Option<String>,
    drop_table: bool,
    insert_ignore: bool,
    table_comment: Option<String>,
    not_null: bool,
    empty_as_null: String,

    varchar_max: i64,
    text_max: i64,
    max_ident: usize,
    merge_scan: String,
    merge_scan_limit: i64,
    fill_down: bool,
    python: Option<String>,
    print_errors: usize,
}

impl Default for Opts {
    fn default() -> Self {
        Opts {
            path: String::new(), sheet: String::new(), table: String::new(),
            header_row: 1, data_row: 2, rows: None,
            out: None, err_file: None, report: None,
            force: false, hints: true, progress: true, color: true,
            scan_only: false,
            batch_size: 200, sample_ratio: 0.10, sample_threshold: 2000,
            sample_min: 1000, sample_max: 20000, full_scan: false, seed: None,
            add_id: None, primary_key: None, drop_table: false, insert_ignore: false,
            table_comment: None, not_null: false, empty_as_null: "auto".into(),
            varchar_max: 1000, text_max: 16000, max_ident: 64,
            merge_scan: "auto".into(), merge_scan_limit: 64, fill_down: true,
            python: None, print_errors: 20,
        }
    }
}

const HELP_ZH: &str = r#"xlsxtomysql —— Excel 转 MySQL 建表 + 插入语句（Rust 单文件版）

用法:
  xlsxtomysql 文件名.xlsx  sheet名  新表名  字段名称所在行  第一个数据所在行  [共几行]

  文件名.xlsx      支持 Excel 2007+ (.xlsx) 与 Excel 2003 (.xls)
  sheet名          工作表名
  新表名           生成的表名，同时作为输出文件名 <新表名>.sql
  字段名称所在行   表头所在行号（从 1 开始）
  第一个数据所在行 第一条数据所在行号
  共几行           可选，默认直到文件尾

主要选项:
  --out PATH              输出 sql 路径（默认 源文件同目录/<新表名>.sql）
  --err-file PATH         失败行文件（默认 源文件同目录/errrows.xlsx）
  --report PATH           把控制台报表另存一份文本
  --force                 发现非标准格式时仍然继续转换
  --scan-only             只扫描并输出类型预览，不生成 sql
  --no-hints              关闭字段名语义推断（手机号/金额等）
  --hints                 重新打开（默认就是开的，主要用来覆盖前面的 --no-hints）
  --full-scan             强制全量扫描（大文件会很慢）
  --sample-ratio F        抽样比例，默认 0.10
  --sample-threshold N    超过多少行启用抽样跳跃扫描，默认 2000
  --sample-min N          抽样行数下限，默认 1000
  --sample-max N          抽样行数上限，默认 20000
  --seed N                抽样起点随机种子（便于复现）
  --batch-size N          每条 INSERT 的行数，默认 200
  --empty-as-null M       auto | always | never，默认 auto
  --varchar-max N         varchar 上限，默认 1000
  --text-max N            text 上限，超过改用 longtext，默认 16000
  --max-ident N           字段名最大长度，默认 64
  --merge-scan M          auto | on | off，默认 auto
  --merge-scan-limit N    合并扫描的最大文件 MB，默认 64
  --no-fill-down          不自动补齐数据区合并单元格
  --add-id NAME           追加一个自增主键列
  --primary-key COL       指定主键列
  --drop-table            生成 DROP TABLE IF EXISTS
  --insert-ignore         生成 INSERT IGNORE
  --table-comment S       表注释
  --not-null              所有字段加 NOT NULL
  --print-errors N        控制台最多显示多少条失败行，默认 20
  --progress off          关闭进度条（--no-progress 等价）
  --no-color              关闭彩色输出（管道输出自动关闭）
  --python PATH           指定读取 Excel 用的 Python 解释器
  --dump-python           打印内嵌的 Python 助手源码后退出
  --man                   显示详细手册（参数细节、类型推断规则、退出码、已知限制）
  --lang zh|en            强制使用中文/英文输出（默认按系统环境自动判断）
  -h, --help              显示本帮助
  -V, --version           显示版本

语言
  默认跟随系统环境：中文环境（zh）输出中文，其它环境输出英文。
  判定顺序：--lang 参数 > 环境变量 XLSXTOMYSQL_LANG > LC_ALL > LC_MESSAGES > LANG；
  全都没有时：Unix 按英文，Windows 读系统界面语言。
  生成的 sql 里的注释也跟着这个语言走。
"#;

struct ArgStream {
    v: Vec<String>,
    i: usize,
}

impl ArgStream {
    fn value(&mut self, name: &str, inline: Option<String>) -> Result<String, CliError> {
        if let Some(x) = inline {
            return Ok(x);
        }
        self.i += 1;
        self.v.get(self.i).cloned()
            .ok_or_else(|| CliError::usage(format!("--{} 后面需要一个值", name)))
    }
}

fn parse_args(argv: Vec<String>) -> Result<Opts, CliError> {
    let mut o = Opts::default();
    let mut pos: Vec<String> = Vec::new();
    let mut st = ArgStream { v: argv, i: 0 };

    while st.i < st.v.len() {
        let a = st.v[st.i].clone();
        if a == "-h" || a == "--help" {
            print!("{}", help_text());
            std::process::exit(EXIT_OK);
        }
        if a == "-V" || a == "--version" {
            println!("xlsxtomysql {} ({}-{})", VERSION, std::env::consts::OS,
                     std::env::consts::ARCH);
            std::process::exit(EXIT_OK);
        }
        if let Some(rest) = a.strip_prefix("--") {
            let (name, inline) = match rest.split_once('=') {
                Some((n, v)) => (n.to_string(), Some(v.to_string())),
                None => (rest.to_string(), None),
            };
            let num = |s: String, what: &str| -> Result<i64, CliError> {
                s.trim().parse::<i64>()
                    .map_err(|_| CliError::usage(format!("--{} 需要整数", what)))
            };
            match name.as_str() {
                "lang" => {
                    // 语言已在 main 里定下来（必须先于任何输出），这里只吃掉取值
                    let _ = st.value("lang", inline)?;
                }
                "man" => {
                    print!("{}", man_text());
                    std::process::exit(EXIT_OK);
                }
                "out" => o.out = Some(st.value("out", inline)?),
                "err-file" => o.err_file = Some(st.value("err-file", inline)?),
                "report" => o.report = Some(st.value("report", inline)?),
                "force" => o.force = true,
                "hints" => o.hints = true,
                "no-hints" => o.hints = false,
                "scan-only" => o.scan_only = true,
                "dump-python" => {
                    // 必须在位置参数校验之前处理：只导出源码，不需要文件参数
                    print!("{}", PY_HELPER);
                    std::process::exit(EXIT_OK);
                }
                "no-color" => o.color = false,
                "progress" => {
                    let v = st.value("progress", inline)?;
                    o.progress = !matches!(v.as_str(), "off" | "no" | "false" | "0");
                }
                "no-progress" => o.progress = false,
                "full-scan" => o.full_scan = true,
                "drop-table" => o.drop_table = true,
                "insert-ignore" => o.insert_ignore = true,
                "not-null" => o.not_null = true,
                "add-id" => o.add_id = Some(st.value("add-id", inline)?),
                "primary-key" => o.primary_key = Some(st.value("primary-key", inline)?),
                "table-comment" => o.table_comment = Some(st.value("table-comment", inline)?),
                "empty-as-null" => {
                    let v = st.value("empty-as-null", inline)?;
                    if !["auto", "always", "never"].contains(&v.as_str()) {
                        return Err(CliError::usage("--empty-as-null 只能是 auto / always / never"));
                    }
                    o.empty_as_null = v;
                }
                "batch-size" => {
                    o.batch_size = num(st.value("batch-size", inline)?, "batch-size")?.max(1);
                }
                "sample-ratio" => {
                    o.sample_ratio = st.value("sample-ratio", inline)?.trim().parse()
                        .map_err(|_| CliError::usage("--sample-ratio 需要小数"))?;
                }
                "sample-threshold" => {
                    o.sample_threshold = num(st.value("sample-threshold", inline)?, "sample-threshold")?;
                }
                "sample-min" => o.sample_min = num(st.value("sample-min", inline)?, "sample-min")?,
                "sample-max" => o.sample_max = num(st.value("sample-max", inline)?, "sample-max")?,
                "seed" => {
                    o.seed = Some(num(st.value("seed", inline)?, "seed")? as u64);
                }
                "varchar-max" => o.varchar_max = num(st.value("varchar-max", inline)?, "varchar-max")?,
                "text-max" => o.text_max = num(st.value("text-max", inline)?, "text-max")?,
                "max-ident" => {
                    o.max_ident = num(st.value("max-ident", inline)?, "max-ident")?.max(1) as usize;
                }
                "merge-scan" => {
                    let v = st.value("merge-scan", inline)?;
                    if !["auto", "on", "off"].contains(&v.as_str()) {
                        return Err(CliError::usage("--merge-scan 只能是 auto / on / off"));
                    }
                    o.merge_scan = v;
                }
                "merge-scan-limit" => {
                    o.merge_scan_limit = num(st.value("merge-scan-limit", inline)?, "merge-scan-limit")?;
                }
                "no-fill-down" => o.fill_down = false,
                "python" => o.python = Some(st.value("python", inline)?),
                "print-errors" => {
                    o.print_errors = num(st.value("print-errors", inline)?, "print-errors")?.max(0) as usize;
                }
                _ => return Err(CliError::usage(format!("未知选项 --{}（详见 --help）", name))),
            }
        } else if a.starts_with('-') && a.len() > 1 {
            return Err(CliError::usage(format!("未知选项 {}（详见 --help）", a)));
        } else {
            pos.push(a);
        }
        st.i += 1;
    }

    if pos.len() < 5 {
        return Err(CliError::usage(format!(
            "参数不足：需要 文件名 sheet名 新表名 字段名称所在行 第一个数据所在行 [共几行]，实际给了 {} 个\n详见 --help",
            pos.len())));
    }
    if pos.len() > 6 {
        return Err(CliError::usage(format!("参数过多（{} 个），最后一个是可选的「共几行」", pos.len())));
    }
    o.path = pos[0].clone();
    o.sheet = pos[1].clone();
    o.table = pos[2].clone();
    o.header_row = pos[3].parse()
        .map_err(|_| CliError::usage("「字段名称所在行」需要是正整数"))?;
    o.data_row = pos[4].parse()
        .map_err(|_| CliError::usage("「第一个数据所在行」需要是正整数"))?;
    if pos.len() == 6 {
        let n: i64 = pos[5].parse()
            .map_err(|_| CliError::usage("「共几行」需要是非负整数"))?;
        if n < 0 {
            return Err(CliError::usage("「共几行」不能是负数"));
        }
        if n > 0 {
            o.rows = Some(n);
        }
    }
    if o.header_row < 1 {
        return Err(CliError::usage("「字段名称所在行」从 1 开始，不能小于 1"));
    }
    if o.data_row < 1 {
        return Err(CliError::usage("「第一个数据所在行」从 1 开始，不能小于 1"));
    }
    if o.data_row <= o.header_row {
        return Err(CliError::usage(format!(
            "「第一个数据所在行」({}) 必须大于「字段名称所在行」({})", o.data_row, o.header_row)));
    }
    Ok(o)
}


// ===========================================================================
// 与内嵌 Python 助手的交互
// ===========================================================================
/// 把若干段拼成一个平台原生路径（Windows 用 \，其它用 /）。
fn join_path(base: &str, parts: &[&str]) -> String {
    let mut p = PathBuf::from(base);
    for x in parts {
        p.push(x);
    }
    p.to_string_lossy().into_owned()
}

fn env_first(keys: &[&str]) -> Option<String> {
    for k in keys {
        if let Ok(v) = std::env::var(k) {
            if !v.trim().is_empty() {
                return Some(v);
            }
        }
    }
    None
}

/// 在 PATH 里找一个可执行文件。
/// Windows 上裸名字（python）通常不是可执行文件，真正的文件是 python.exe，
/// 所以要按 PATHEXT 逐个补扩展名；Unix 上不加扩展名。
fn which(name: &str) -> Option<String> {
    if name.contains('/') || name.contains('\\') {
        return if Path::new(name).is_file() { Some(name.to_string()) } else { None };
    }
    let exts: Vec<String> = if cfg!(windows) {
        let raw = std::env::var("PATHEXT").unwrap_or_else(|_| ".COM;.EXE;.BAT;.CMD".into());
        let mut v: Vec<String> = vec![String::new()];
        for e in raw.split(';') {
            let e = e.trim().to_lowercase();
            if !e.is_empty() {
                v.push(e);
            }
        }
        v
    } else {
        vec![String::new()]
    };
    let path = std::env::var_os("PATH")?;
    for dir in std::env::split_paths(&path) {
        if dir.as_os_str().is_empty() {
            continue;
        }
        for ext in &exts {
            let cand = dir.join(format!("{}{}", name, ext));
            if cand.is_file() {
                return Some(cand.to_string_lossy().into_owned());
            }
        }
    }
    None
}

/// 找一个能读 Excel 的 Python 解释器。顺序：
///   --python  >  $XLSXTOMYSQL_PYTHON  >  本工具自带的 venv  >  PATH  >  平台常见位置。
/// 找不到也不报错：真正用到时助手会给出带提示的错误。
fn find_python(opts: &Opts) -> String {
    if let Some(p) = opts.python.as_ref() {
        return p.clone();
    }
    if let Some(p) = env_first(&["XLSXTOMYSQL_PYTHON"]) {
        return p;
    }
    let mut cands: Vec<String> = Vec::new();
    // 先看二进制自己旁边的 venv：发布包（tar.gz / zip）解压或装到任意目录时，
    // install.sh / install.ps1 建的 venv 就在 <可执行文件目录>/../share/xlsxtomysql/venv，
    // 这样便携安装不用再设任何环境变量。
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            let rel: [&[&str]; 2] = if cfg!(windows) {
                [&["..", "share", "xlsxtomysql", "venv", "Scripts", "python.exe"],
                 &["..", "..", "share", "xlsxtomysql", "venv", "Scripts", "python.exe"]]
            } else {
                [&["..", "share", "xlsxtomysql", "venv", "bin", "python"],
                 &["..", "..", "share", "xlsxtomysql", "venv", "bin", "python"]]
            };
            for parts in rel {
                let mut p = dir.to_path_buf();
                for seg in parts {
                    p.push(seg);
                }
                cands.push(p.to_string_lossy().into_owned());
            }
        }
    }
    if let Some(home) = env_first(&["HOME", "USERPROFILE"]) {
        if cfg!(windows) {
            cands.push(join_path(&home, &["AppData", "Local", "xlsxtomysql", "venv", "Scripts", "python.exe"]));
            cands.push(join_path(&home, &[".local", "share", "xlsxtomysql", "venv", "Scripts", "python.exe"]));
        } else {
            cands.push(join_path(&home, &[".local", "share", "xlsxtomysql", "venv", "bin", "python"]));
            cands.push(join_path(&home, &[".local", "bin", "python3"]));
        }
    }
    if cfg!(windows) {
        if let Some(la) = env_first(&["LOCALAPPDATA"]) {
            cands.push(join_path(&la, &["xlsxtomysql", "venv", "Scripts", "python.exe"]));
        }
    } else {
        for p in ["/usr/bin/python3", "/usr/local/bin/python3", "/opt/homebrew/bin/python3"] {
            cands.push(p.to_string());
        }
    }
    for c in &cands {
        if Path::new(c).is_file() {
            return c.clone();
        }
    }
    // PATH 上的解释器：Unix 优先 python3；Windows 优先 python（另有 py 启动器兜底）
    let names: &[&str] = if cfg!(windows) {
        &["python", "python3", "py"]
    } else {
        &["python3", "python"]
    };
    for name in names {
        if let Some(p) = which(name) {
            return p;
        }
    }
    if cfg!(windows) { "python".to_string() } else { "python3".to_string() }
}

/// 是否改用「临时 .py 文件」的方式启动助手。
/// Windows 的整条命令行上限约 32K（CreateProcess 限制），而内嵌助手源码有 35K，
/// 用 `python -c <源码>` 一启动就会被拒（报错通常是「文件名或扩展名太长」）。
/// 所以 Windows 上落一份临时脚本再执行；Unix 的 ARG_MAX 有 2MB，继续用 -c 不落盘。
/// 想在本机验证这条路径：XLSXTOMYSQL_PYFILE=1。
fn use_helper_file() -> bool {
    if let Ok(v) = std::env::var("XLSXTOMYSQL_PYFILE") {
        let v = v.trim();
        if !v.is_empty() && v != "0" {
            return true;
        }
    }
    cfg!(windows) && PY_HELPER.len() > 24_000
}

/// 把助手源码写到临时文件（文件名带源码长度与滚动校验，同一版本只写一次）。
fn helper_script_path() -> io::Result<String> {
    let mut dir = std::env::temp_dir();
    dir.push("xlsxtomysql");
    if !dir.is_dir() {
        std::fs::create_dir_all(&dir)?;
    }
    let sum = PY_HELPER.bytes().fold(0u32, |a, b| a.wrapping_mul(31).wrapping_add(b as u32));
    let mut p = dir;
    p.push(format!("py_helper_{:08x}_{}.py", sum, PY_HELPER.len()));
    if !p.is_file() {
        std::fs::write(&p, PY_HELPER)?;
    }
    Ok(p.to_string_lossy().into_owned())
}

fn helper_command(py: &str) -> Command {
    let mut c = Command::new(py);
    let mut launched = false;
    if use_helper_file() {
        // 落盘失败（只读 %TEMP%）就退回命令行，交给上层报错
        if let Ok(script) = helper_script_path() {
            c.arg(&script);
            launched = true;
        }
    }
    if !launched {
        c.arg("-c").arg(PY_HELPER);
    }
    c.env("XLSXTOMYSQL_LANG", if lang_zh() { "zh" } else { "en" })
        .env("PYTHONIOENCODING", "utf-8")
        .env("PYTHONDONTWRITEBYTECODE", "1")
        .env("PYTHONUTF8", "1");
    c
}

/// 一次性调用（probe / errrows）：返回 (退出码, stdout, stderr)
fn run_helper_once(py: &str, args: &[String], stdin_data: Option<&str>) -> Result<(i32, String, String), CliError> {
    let mut cmd = helper_command(py);
    cmd.args(args)
        .stdin(if stdin_data.is_some() { Stdio::piped() } else { Stdio::null() })
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    let mut child = cmd.spawn().map_err(|e| {
        CliError::general(format!(
            "无法启动 Python 解释器 {}\n  {}\n  可用 --python 指定一个装了 openpyxl 的解释器",
            py, e))
    })?;
    if let Some(data) = stdin_data {
        if let Some(mut si) = child.stdin.take() {
            let _ = si.write_all(data.as_bytes());
        }
    }
    let output = child.wait_with_output()
        .map_err(|e| CliError::general(format!("Python 助手异常: {}", e)))?;
    Ok((
        output.status.code().unwrap_or(-1),
        String::from_utf8_lossy(&output.stdout).into_owned(),
        String::from_utf8_lossy(&output.stderr).into_owned(),
    ))
}

/// 从 #FATAL<tab>退出码<tab>消息 中取回 (退出码, 消息)
fn fatal_from_stdout(stdout: &str) -> Option<(i32, String)> {
    stdout.lines().find_map(|l| {
        let rest = l.strip_prefix("#FATAL\t")?;
        let (code, msg) = rest.split_once('\t')?;
        Some((code.trim().parse().unwrap_or(EXIT_ERROR), unescape(msg)))
    })
}

fn stderr_tail(stderr: &str, n: usize) -> String {
    let all: Vec<&str> = stderr.lines().filter(|l| !l.trim().is_empty()).collect();
    let start = all.len().saturating_sub(n);
    all[start..].join("\n")
}

// ---- 事件 ----
// 只保留 Rust 侧真正要用的载荷：工作簿信息与表头本来就走 probe 一次取全，
// 所以 #SHEET / #END / #OK / 普通输出行这些事件无需带数据，仅用于占位与排错。
enum Ev {
    Kind(String),
    SheetInfo(i64, i64),
    Sheets(Vec<String>),
    Merges(Vec<(i64, i64, i64, i64)>),
    Formula(bool),
    Hidden(Vec<i64>, Vec<i64>),
    EmbedImg(bool),
    Hfill(Vec<(usize, String)>),
    Pct(i64),
    Tz(Vec<(usize, i64)>),
    Header(Vec<Cell>),
    Row(i64, Vec<Cell>),
    End,
    Ok,
    Fatal(i32, String),
    Other,
}

fn parse_event(line: &str) -> Ev {
    let mut it = line.split('\t');
    let tag = it.next().unwrap_or("");
    match tag {
        "#KIND" => Ev::Kind(it.next().unwrap_or("").to_string()),
        "#SHEET" => {
            let _name = unescape(it.next().unwrap_or(""));
            let a = it.next().unwrap_or("0").parse().unwrap_or(0);
            let b = it.next().unwrap_or("0").parse().unwrap_or(0);
            Ev::SheetInfo(a, b)
        }
        "#SHEETS" => Ev::Sheets(it.map(unescape).collect()),
        "#MERGES" => {
            let raw = it.next().unwrap_or("");
            let mut v = Vec::new();
            for part in raw.split('|').filter(|s| !s.is_empty()) {
                let nums: Vec<i64> = part.split(',').filter_map(|x| x.parse().ok()).collect();
                if nums.len() == 4 {
                    v.push((nums[0], nums[1], nums[2], nums[3]));
                }
            }
            Ev::Merges(v)
        }
        "#FORMULA" => Ev::Formula(it.next().unwrap_or("0") == "1"),
        "#HIDDEN" => {
            let cols = it.next().unwrap_or("");
            let rows_ = it.next().unwrap_or("");
            let parse_list = |s: &str| -> Vec<i64> {
                s.split(',').filter(|x| !x.is_empty()).filter_map(|x| x.parse().ok()).collect()
            };
            Ev::Hidden(parse_list(cols), parse_list(rows_))
        }
        "#EMBEDIMG" => Ev::EmbedImg(it.next().unwrap_or("0") == "1"),
        "#HFILL" => {
            // 扁平成对的 列号 / 值
            let items: Vec<&str> = it.collect();
            let mut v = Vec::new();
            let mut i = 0;
            while i + 1 < items.len() {
                if let Ok(c) = items[i].parse::<usize>() {
                    v.push((c, unescape(&items[i + 1])));
                }
                i += 2;
            }
            Ev::Hfill(v)
        }
        "#PCT" => Ev::Pct(it.next().unwrap_or("0").parse().unwrap_or(0)),
        "#TZ" => {
            // 扁平成对的 列号 / 个数
            let items: Vec<&str> = it.collect();
            let mut v = Vec::new();
            let mut i = 0;
            while i + 1 < items.len() {
                if let (Ok(c), Ok(n)) = (items[i].parse::<usize>(), items[i + 1].parse::<i64>()) {
                    v.push((c, n));
                }
                i += 2;
            }
            Ev::Tz(v)
        }
        "#HEADER" => Ev::Header(it.map(parse_cell).collect()),
        "#ROW" => {
            let n = it.next().unwrap_or("0").parse().unwrap_or(0);
            Ev::Row(n, it.map(parse_cell).collect())
        }
        "#END" => {
            // #END 后面跟的总行数/实际行数/读取行数对 Rust 侧无用（改用流里
            // 实际见到的行号推算），这里只做消费。
            for _ in 0..3 {
                it.next();
            }
            Ev::End
        }
        "#OK" => Ev::Ok,
        "#FATAL" => {
            // 协议：#FATAL<tab>退出码<tab>消息
            let code = it.next().unwrap_or("1").parse().unwrap_or(EXIT_ERROR);
            Ev::Fatal(code, tr(&unescape(it.next().unwrap_or(""))))
        }
        _ => Ev::Other,
    }
}

/// Python 助手的一次流式运行：两个线程分别接管 stdout（数据）与 stderr（进度/日志）
struct Stream {
    rx: Receiver<Ev>,
    child: Child,
}

impl Stream {
    fn start(py: &str, args: &[String], label: &str, show_progress: bool) -> Result<Stream, CliError> {
        let mut cmd = helper_command(py);
        cmd.args(args).stdin(Stdio::null()).stdout(Stdio::piped()).stderr(Stdio::piped());
        let mut child = cmd.spawn().map_err(|e| {
            CliError::general(format!("无法启动 Python 解释器 {}\n  {}", py, e))
        })?;

        let (tx, rx) = channel::<Ev>();
        let stdout = child.stdout.take().expect("stdout 已设为管道");
        thread::spawn(move || {
            let reader = BufReader::with_capacity(1 << 20, stdout);
            for line in reader.lines() {
                let line = match line {
                    Ok(l) => l,
                    Err(_) => break,
                };
                if line.is_empty() {
                    continue;
                }
                if tx.send(parse_event(&line)).is_err() {
                    break;
                }
            }
        });

        let stderr = child.stderr.take().expect("stderr 已设为管道");
        let label_owned = label.to_string();
        thread::spawn(move || {
            let reader = BufReader::new(stderr);
            let mut bar: Option<Bar> = None;
            for line in reader.lines() {
                let line = match line {
                    Ok(l) => l,
                    Err(_) => break,
                };
                if let Some(rest) = line.strip_prefix("#P\t") {
                    if !show_progress {
                        continue;
                    }
                    let mut parts = rest.split('\t');
                    let n: i64 = parts.next().unwrap_or("0").parse().unwrap_or(0);
                    let total: i64 = parts.next().unwrap_or("0").parse().unwrap_or(0);
                    let b = bar.get_or_insert_with(|| Bar::new(&label_owned, true));
                    b.set_total(total);
                    b.set(n);
                } else if !line.trim().is_empty() {
                    eprintln!("  ▸ {}", tr(&line));
                }
            }
            if let Some(mut b) = bar {
                b.finish("");
            }
        });

        Ok(Stream { rx, child })
    }

    fn kill(&mut self) {
        let _ = self.child.kill();
        let _ = self.child.wait();
    }
}


// ===========================================================================
// 字段名清洗
// ===========================================================================
fn sanitize_ident(raw: &str, fallback: &str, maxlen: usize) -> (String, String) {
    let mut notes: Vec<String> = Vec::new();
    let cleaned: String = raw
        .replace('\u{3000}', " ")
        .replace('\u{a0}', " ")
        .chars()
        .filter(|c| !matches!(c, '\u{0}'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{7f}'))
        .map(|c| if c == '\n' || c == '\r' { ' ' } else { c })
        .collect();
    let s = cleaned.trim().to_string();
    if s.is_empty() {
        return (fallback.to_string(), t("原字段名为空"));
    }
    if s.chars().any(|c| c.is_whitespace()) {
        notes.push(t("含空白"));
    }
    // 非字母数字下划线（含中文保留）→ 下划线
    let mut collapsed = String::new();
    let mut prev_us = false;
    for c in s.chars() {
        let keep = c.is_alphanumeric() || c == '_';
        if !keep {
            if !prev_us {
                collapsed.push('_');
            }
            prev_us = true;
        } else {
            collapsed.push(c);
            prev_us = false;
        }
    }
    let mut s2 = collapsed.trim_matches('_').to_string();
    if s2 != s {
        notes.push(t("含非法字符"));
    }
    if s2.is_empty() {
        return (fallback.to_string(), t("原字段名全为非法字符"));
    }
    if s2.chars().next().map(|c| c.is_ascii_digit()).unwrap_or(false) {
        s2 = format!("c_{}", s2);
        notes.push(t("以数字开头"));
    }
    if s2.chars().count() > maxlen {
        s2 = s2.chars().take(maxlen).collect();
        notes.push(t("超长截断"));
    }
    let note = notes.join(sep());
    (s2, note)
}

#[derive(Clone)]
struct Col {
    orig: String,
    name: String,
    note: String,
    mysql_type: String,
    reason: String,
}

fn col_name(idx: usize) -> String {
    let mut s = String::new();
    let mut n = idx;
    while n > 0 {
        let r = (n - 1) % 26;
        s.insert(0, (b'A' + r as u8) as char);
        n = (n - 1) / 26;
    }
    s
}

fn last_non_empty(cells: &[Cell]) -> usize {
    let mut w = 0usize;
    for (i, c) in cells.iter().enumerate() {
        if !c.is_empty() {
            w = i + 1;
        }
    }
    w
}

/// 合并单元格里只有左上角有值，这里把值补到整个合并区：
///   * 字段名行落在纵向合并区时，值在更上面那行（由 helper 通过 #HFILL 给出）
///   * 横向合并则把左上角的值向右复制
fn fill_merged(merges: &[(i64, i64, i64, i64)], header_row: i64, header: &mut Vec<Cell>,
               hfill: &[(usize, String)]) -> (usize, usize) {
    // 纵向合并：值在合并区左上角那一行，helper 已经把值取出来给到 hfill
    for (c, v) in hfill {
        if *c == 0 {
            continue;
        }
        while header.len() < *c {
            header.push(EMPTY_CELL);
        }
        if header[*c - 1].is_empty() {
            header[*c - 1] = Cell::Text(v.clone());
        }
    }
    let mut touched = 0usize;
    let mut vertical = 0usize;
    for (r1, c1, r2, c2) in merges {
        if !(*r1 <= header_row && header_row <= *r2) {
            continue;
        }
        let ci = (*c1 - 1).max(0) as usize;
        if ci >= header.len() {
            continue;
        }
        let base = header[ci].clone();
        if base.is_empty() {
            continue;
        }
        let hi = (*c2).max(0) as usize;
        while header.len() < hi {
            header.push(EMPTY_CELL);
        }
        let mut added = 0usize;
        for c in ci..hi {
            if header[c].is_empty() {
                header[c] = base.clone();
                added += 1;
            }
        }
        if *r2 > *r1 {
            vertical += 1;
        }
        if added > 0 {
            touched += 1;
        }
    }
    (touched, vertical)
}

/// 数据区合并单元格填充：一个值跨多行（如一个学生占 3 行）时，把左上角的内容补齐整块
struct MergeFiller {
    regions: Vec<(i64, i64, i64, i64)>,
    cache: HashMap<(i64, i64), Cell>,
    filled: i64,
}

impl MergeFiller {
    fn new(merges: &[(i64, i64, i64, i64)], data_row: i64) -> Self {
        let regions = merges.iter().cloned()
            .filter(|(r1, c1, r2, c2)| *r2 >= data_row && (*r2 > *r1 || *c2 > *c1))
            .collect();
        MergeFiller { regions, cache: HashMap::new(), filled: 0 }
    }

    fn apply(&mut self, row_no: i64, cells: &mut Vec<Cell>) {
        if self.regions.is_empty() {
            return;
        }
        for i in 0..self.regions.len() {
            let (r1, c1, _, _) = self.regions[i];
            if r1 != row_no {
                continue;
            }
            let ci = (c1 - 1) as usize;
            if ci < cells.len() && !cells[ci].is_empty() {
                self.cache.insert((r1, c1), cells[ci].clone());
            }
        }
        for i in 0..self.regions.len() {
            let (r1, c1, r2, c2) = self.regions[i];
            if !(r1 <= row_no && row_no <= r2) {
                continue;
            }
            let v = match self.cache.get(&(r1, c1)) {
                Some(x) => x.clone(),
                None => continue,
            };
            while (cells.len() as i64) < c2 {
                cells.push(EMPTY_CELL);
            }
            let hi = c2.min(cells.len() as i64);
            for ci in (c1 - 1)..hi {
                let j = ci as usize;
                if cells[j].is_empty() {
                    cells[j] = v.clone();
                    self.filled += 1;
                }
            }
        }
    }
}

fn build_columns(header: &[Cell], opts: &Opts) -> Vec<Col> {
    let mut cols = Vec::with_capacity(header.len());
    let mut used: HashSet<String> = HashSet::new();
    for (i, c) in header.iter().enumerate() {
        let head = c.text().trim().to_string();
        let fallback = format!("col_{}", col_name(i + 1).to_lowercase());
        let (base, mut note) = sanitize_ident(&head, &fallback, opts.max_ident);
        if head.is_empty() {
            note = t("原字段名为空");
        }
        let mut name = base.clone();
        let mut k = 1;
        while used.contains(&name) {
            k += 1;
            name = format!("{}_{}", base, k);
        }
        let mut parts: Vec<String> = Vec::new();
        if !note.is_empty() {
            parts.push(note);
        }
        if name != base {
            parts.push(tf("与前面的字段重名(已改名 {})", &[&name]));
        }
        used.insert(name.clone());
        if RESERVED_HINTS.contains(&name.to_lowercase().as_str()) {
            parts.push(t("与 SQL 关键字同名(已用反引号包裹)"));
        }
        cols.push(Col {
            orig: head,
            name,
            note: parts.join(sep()),
            mysql_type: "varchar(255)".into(),
            reason: String::new(),
        });
    }
    cols
}

// ===========================================================================
// 类型推断
// ===========================================================================
#[derive(Default)]
struct ColAcc {
    total_cells: i64,
    non_empty: i64,
    error_cells: i64,
    kinds: HashMap<Kind, i64>,
    min_int: Option<i128>,
    max_int: Option<i128>,
    max_scale: i64,
    max_int_digits: i64,
    max_len: i64,
    max_bytes: i64,
    out_of_range: i64,     // 超出 MySQL 日期/时间范围的值
    huge_numbers: i64,     // >= 1e15 的数值，Excel 本身已丢精度
}

impl ColAcc {
    fn add(&mut self, c: &Cell) {
        self.total_cells += 1;
        match c.kind() {
            Kind::Empty => return,
            Kind::Err => {
                self.error_cells += 1;
                return;
            }
            _ => {}
        }
        // Excel 只有 15 位有效数字，更长的大数在保存时就已经丢精度了
        if let Cell::Dec(f) | Cell::Sci(f) = c {
            if f.abs() >= 1e15 {
                self.huge_numbers += 1;
            }
        }
        self.non_empty += 1;
        *self.kinds.entry(c.kind()).or_insert(0) += 1;

        match c {
            Cell::Int(v) => {
                self.min_int = Some(self.min_int.map_or(*v, |x| x.min(*v)));
                self.max_int = Some(self.max_int.map_or(*v, |x| x.max(*v)));
                self.max_int_digits = self.max_int_digits.max(v.abs().to_string().len() as i64);
            }
            Cell::Dec(f) | Cell::Sci(f) => {
                let s = fmt_float(*f);
                let body = s.trim_start_matches(['+', '-']).to_string();
                match body.split_once('.') {
                    Some((ip, fp)) => {
                        let fp = fp.split(['e', 'E']).next().unwrap_or("");
                        let ip = ip.split(['e', 'E']).next().unwrap_or("");
                        let il = if ip.is_empty() { 1 } else { ip.len() as i64 };
                        self.max_int_digits = self.max_int_digits.max(il);
                        self.max_scale = self.max_scale.max(fp.len() as i64);
                    }
                    None => {
                        if body.contains('e') || body.contains('E') {
                            self.max_int_digits = self.max_int_digits.max(1);
                        } else {
                            self.max_int_digits = self.max_int_digits.max(body.len() as i64);
                        }
                    }
                }
            }
            Cell::Text(s) | Cell::Oor(s) => {
                if matches!(c, Cell::Oor(_)) {
                    self.out_of_range += 1;
                }
                self.max_len = self.max_len.max(s.chars().count() as i64);
                self.max_bytes = self.max_bytes.max(s.len() as i64);
            }
            Cell::Date(_) | Cell::Time(_) | Cell::DateTime(_) => {
                self.max_len = self.max_len.max(19);
                self.max_bytes = self.max_bytes.max(19);
            }
            _ => {}
        }
    }
}

fn nice_varchar(n: i64, cap: i64) -> i64 {
    for size in VARCHAR_LADDER {
        if n <= *size {
            return (*size).min(cap);
        }
    }
    cap
}

fn resolve_type(acc: &ColAcc, name: &str, opts: &Opts) -> (String, String) {
    let kinds: Vec<Kind> = acc.kinds.iter().filter(|(_, v)| **v > 0).map(|(k, _)| *k).collect();
    if kinds.is_empty() {
        // 整句一次成型：中英文的括号/逗号位置不同，拼接片段会拼出病句
        let extra = if acc.error_cells > 0 {
            tf("（该列在样本中全为空，另有 {} 个 Excel 错误值）", &[&acc.error_cells])
        } else {
            t("（该列在样本中全为空）")
        };
        return ("varchar(255)".into(), tf("无有效数据，默认 varchar(255){}", &[&extra]));
    }

    let numeric_kinds = [Kind::Int, Kind::Dec, Kind::Sci];
    let pure_numeric = kinds.iter().all(|k| numeric_kinds.contains(k));
    let low = name.to_lowercase();
    let id_hint = ID_HINTS.iter().any(|h| low.contains(h));
    let money_hint = MONEY_HINTS.iter().any(|h| low.contains(h));

    if opts.hints && id_hint && pure_numeric {
        let n = nice_varchar(acc.max_len.max(18), opts.varchar_max).max(18);
        return (format!("varchar({})", n),
                format!("字段名疑似标识类(手机/证件/账号)，按文本保存，样本最长 {} 字符", acc.max_len));
    }

    if pure_numeric {
        let is_int_only = kinds.iter().all(|k| *k == Kind::Int) && acc.max_scale == 0;
        if is_int_only {
            let (lo, hi) = match (acc.min_int, acc.max_int) {
                (Some(a), Some(b)) => (a, b),
                _ => return ("int".into(), "全部为整数".into()),
            };
            for (tname, tlo, thi) in INT_RANGES {
                if *tlo <= lo && hi <= *thi {
                    if money_hint && opts.hints {
                        let p = 14i64.max(acc.max_int_digits + 2);
                        return (format!("decimal({},2)", p),
                                format!("字段名疑似金额类，样本范围 {}~{}", lo, hi));
                    }
                    return ((*tname).to_string(), format!("全部为整数，样本范围 {}~{}", lo, hi));
                }
            }
            return ("decimal(20,0)".into(), format!("整数超出 bigint 范围（样本 {}~{}）", lo, hi));
        }
        if kinds.contains(&Kind::Sci) {
            return ("double".into(), "样本含科学计数法数值，使用 double".into());
        }
        let mut p = acc.max_int_digits + acc.max_scale;
        let mut s = acc.max_scale;
        if s == 0 {
            s = 2;
            p = p.max(10);
        }
        if money_hint && opts.hints {
            p = p.max(14);
            s = s.max(2);
        }
        if p > 65 || s > 30 {
            return ("double".into(), format!("精度需求 decimal({},{}) 超出 MySQL 上限，改用 double", p, s));
        }
        p = p.max(s + 1).max(4);
        return (format!("decimal({},{})", p, s),
                format!("含小数的数值，整数位最长 {} 位、小数 {} 位", acc.max_int_digits, acc.max_scale));
    }

    let is_only = |set: &[Kind]| kinds.iter().all(|k| set.contains(k));
    if is_only(&[Kind::Bool]) {
        return ("tinyint(1)".into(), "只有真假值".into());
    }
    if is_only(&[Kind::Date]) {
        return ("date".into(), "全部为日期".into());
    }
    if is_only(&[Kind::Time]) {
        return ("time".into(), "全部为时间".into());
    }
    if is_only(&[Kind::DateTime]) {
        return ("datetime".into(), "全部为日期时间".into());
    }
    if is_only(&[Kind::Date, Kind::DateTime]) {
        return ("datetime".into(), "日期与日期时间混合，统一为 datetime".into());
    }
    if is_only(&[Kind::Date, Kind::Time]) || is_only(&[Kind::Time, Kind::DateTime, Kind::Date]) {
        let n = acc.max_len.max(19);
        return (format!("varchar({})", nice_varchar(n, opts.varchar_max)),
                "日期与时间混杂，按文本保存以免失真".into());
    }

    // 落到文本
    let n = acc.max_len.max(1);
    let only_text = kinds.iter().all(|k| *k == Kind::Text);
    let has_text = kinds.contains(&Kind::Text);
    let mut reason = if acc.out_of_range > 0 {
        format!("含 {} 个超出 MySQL 日期/时间范围的值，整列按文本保存（样本最长 {} 字符）",
                acc.out_of_range, n)
    } else if only_text || (!has_text && !pure_numeric) {
        format!("文本内容，样本最长 {} 字符", n)
    } else {
        format!("包含非数值/非日期内容，降级为文本，样本最长 {} 字符", n)
    };
    if acc.max_len >= 100 && (acc.max_len as f64) > n as f64 * 0.9 {
        reason.push_str("（已留少量余量）");
    }
    if n <= opts.varchar_max {
        return (format!("varchar({})", nice_varchar(n, opts.varchar_max)), reason);
    }
    // MySQL 的 text 上限是 65535 **字节**，中文一个字 3 字节，不能只看字符数
    if n <= opts.text_max && acc.max_bytes <= MYSQL_TEXT_BYTES {
        return ("text".into(),
                format!("内容较长（最长 {} 字符 / {} 字节），使用 text", n, acc.max_bytes));
    }
    ("longtext".into(),
     format!("内容超长（最长 {} 字符 / {} 字节），使用 longtext", n, acc.max_bytes))
}

/// MySQL 的 text 上限（字节）
const MYSQL_TEXT_BYTES: i64 = 65535;

// ===========================================================================
// 格式检查
// ===========================================================================
#[derive(Clone)]
struct Issue {
    level: &'static str,
    row: i64,
    col: Option<usize>,
    title: String,
    detail: String,
}

impl Issue {
    fn where_(&self) -> String {
        match self.col {
            None => format!("第 {} 行", self.row),
            Some(c) => format!("第 {} 行 / 第 {} 列", self.row, col_name(c)),
        }
    }
}

fn is_divider(s: &str) -> bool {
    !s.is_empty() && s.chars().all(|c| matches!(c, '-' | '—' | '_' | '=' | '*' | '·' | '~' | '+' | '#' | '.' | '…') || c.is_whitespace())
}

fn contains_any_ci(s: &str, needles: &[&str]) -> bool {
    let low = s.to_lowercase();
    needles.iter().any(|n| low.contains(n))
}

struct Checker {
    header: Vec<String>,
    width: usize,
    issues: Vec<Issue>,
    empty_rows: i64,
    subtotal_rows: i64,
    seen_data: bool,
    trailing_empty: i64,
    /// 连续空行先记下来：只有后面又出现数据行，才能判定它们是"中间空行"；
    /// 若一直空到结尾，那就是正常的尾部空行（Excel 导出很常见），不该阻断。
    pending_empty: Vec<i64>,
}

impl Checker {
    fn new(header: &[Cell], width: usize) -> Self {
        Checker {
            header: header.iter().map(|c| c.text().trim().to_string()).collect(),
            width,
            issues: Vec::new(),
            empty_rows: 0,
            subtotal_rows: 0,
            seen_data: false,
            trailing_empty: 0,
            pending_empty: Vec::new(),
        }
    }

    /// 检查一行；返回 false 表示这行不该参与类型推断
    fn add(&mut self, row_no: i64, cells: &[Cell]) -> bool {
        let nonempty: Vec<usize> = cells.iter().enumerate()
            .filter(|(_, c)| !c.is_empty()).map(|(i, _)| i).collect();
        let texts: Vec<String> = cells.iter().filter(|c| !c.is_empty())
            .map(|c| c.text().trim().to_string()).collect();
        let tail = nonempty.last().map(|i| i + 1).unwrap_or(0);

        if nonempty.is_empty() {
            self.empty_rows += 1;
            if self.seen_data {
                self.pending_empty.push(row_no);
                self.trailing_empty += 1;
            }
            return false;
        }
        // 前面攒下的连续空行，现在可以确认是数据区中间的空行了
        if !self.pending_empty.is_empty() {
            for rn in std::mem::take(&mut self.pending_empty) {
                self.issues.push(Issue {
                    level: "block", row: rn, col: None,
                    title: "分隔用的空行".into(),
                    detail: "数据区中间出现整行为空的行，会打断数据；请删除该行或加 --force 忽略".into(),
                });
            }
            self.trailing_empty = 0;
        }
        if !self.seen_data && nonempty.len() == 1 {
            let c0 = &cells[nonempty[0]];
            if c0.kind() == Kind::Text {
                let v: String = c0.text().chars().take(20).collect();
                self.issues.push(Issue {
                    level: "warn", row: row_no, col: Some(nonempty[0] + 1),
                    title: "疑似说明行".into(),
                    detail: format!("数据区第一行只有这一列有内容（{}），常见于「单位：元」之类的说明；若确为数据请忽略", v),
                });
            }
        }
        self.seen_data = true;

        if !texts.is_empty()
            && texts.iter().all(|t| is_divider(t))
            && texts.iter().all(|t| t.chars().count() <= 12)
        {
            let preview: Vec<String> = texts.iter().take(4).cloned().collect();
            self.issues.push(Issue {
                level: "block", row: row_no, col: None,
                title: "分隔线/装饰行".into(),
                detail: format!("该行内容像分隔符（{}），不是数据；请删除后重试", preview.join(" | ")),
            });
            return false;
        }

        if !texts.is_empty() && nonempty.len() >= 2 {
            let mut same = 0usize;
            for i in &nonempty {
                if *i < self.header.len() && !self.header[*i].is_empty()
                    && cells[*i].text().trim() == self.header[*i]
                {
                    same += 1;
                }
            }
            let need = (((nonempty.len() as f64) * 0.7) as usize).max(2);
            if same >= need {
                self.issues.push(Issue {
                    level: "block", row: row_no, col: None,
                    title: "重复的字段名行".into(),
                    detail: "该行内容与字段名行相同，疑似重复表头；请删除后重试".into(),
                });
                return false;
            }
        }

        if !texts.is_empty() && nonempty.len() <= 3
            && texts.iter().any(|t| contains_any_ci(t, SUBTOTAL_HINTS))
        {
            self.subtotal_rows += 1;
            let preview: Vec<String> = texts.iter().take(3).cloned().collect();
            self.issues.push(Issue {
                level: "warn", row: row_no, col: None,
                title: "疑似合计/小计行".into(),
                detail: format!("内容: {}；转换时会被当作普通数据行，建议删除", preview.join(" | ")),
            });
        }

        if tail > self.width {
            let extra: Vec<String> = cells[self.width..].iter().filter(|c| !c.is_empty())
                .map(|c| truncate_disp(&c.text(), 12)).collect();
            self.issues.push(Issue {
                level: "block", row: row_no, col: Some(self.width + 1),
                title: "数据超出字段名范围".into(),
                detail: format!("第 {} 列及之后有内容（{}），但字段名行只到 {} 列；请补齐字段名或删除多余列",
                                col_name(self.width + 1), extra.join(" | "), col_name(self.width)),
            });
        }
        true
    }

    fn finalize(&mut self, last_row: i64, first_data_row: i64) {
        if self.trailing_empty > 0 && self.seen_data {
            let at = last_row - self.trailing_empty + 1;
            let n = self.trailing_empty;
            self.issues.push(Issue {
                level: "warn", row: at, col: None,
                title: "末尾空行".into(),
                detail: format!("数据区末尾有 {} 行空白，已自动忽略", n),
            });
        }
        if self.empty_rows > 0 && !self.seen_data {
            self.issues.push(Issue {
                level: "block", row: first_data_row, col: None,
                title: "数据区为空".into(),
                detail: "从「第一个数据所在行」开始没有任何数据，请检查行号参数".into(),
            });
        }
    }
}


// ===========================================================================
// SQL 生成
// ===========================================================================
fn sql_lit(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('\'');
    for c in s.chars() {
        match c {
            '\\' => out.push_str("\\\\"),
            '\'' => out.push_str("\\'"),
            '\0' => out.push_str("\\0"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\x1a' => out.push_str("\\Z"),
            '\t' => out.push_str("\\t"),
            _ => out.push(c),
        }
    }
    out.push('\'');
    out
}

fn ident(name: &str) -> String {
    format!("`{}`", name.replace('`', "``"))
}

fn comment_text(s: &str) -> String {
    s.replace('\\', "\\\\").replace('\'', "''")
}

fn sql_safe(s: &str) -> String {
    s.chars().filter(|c| *c != '\0').collect()
}

#[derive(Clone, Copy)]
enum TypeSpec {
    None,
    Int(i128, i128),
    Dec(i64, i64),
    Str(i64),
}

fn parse_type_spec(t: &str) -> (String, TypeSpec) {
    let base = t.split('(').next().unwrap_or(t).trim().to_lowercase();
    let params: Vec<i64> = t.split_once('(')
        .map(|(_, r)| r.trim_end_matches(')'))
        .unwrap_or("")
        .split(',')
        .filter_map(|x| x.trim().parse::<i64>().ok())
        .collect();
    let spec = match base.as_str() {
        "tinyint" => TypeSpec::Int(-128, 127),
        "smallint" => TypeSpec::Int(-32768, 32767),
        "mediumint" => TypeSpec::Int(-8388608, 8388607),
        "int" | "integer" => TypeSpec::Int(-2147483648, 2147483647),
        "bigint" => TypeSpec::Int(i128::from(i64::MIN), i128::from(i64::MAX)),
        "decimal" | "numeric" => TypeSpec::Dec(
            params.first().copied().unwrap_or(10),
            params.get(1).copied().unwrap_or(0),
        ),
        "varchar" | "char" => TypeSpec::Str(params.first().copied().unwrap_or(255)),
        _ => TypeSpec::None,
    };
    (base, spec)
}

fn is_string_type(base: &str) -> bool {
    matches!(base, "varchar" | "char" | "text" | "mediumtext" | "longtext")
}

// ---- 日期时间 ----
fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m, d)
}

/// 本地时区相对 UTC 的偏移（秒）。拿不到就返回 0（按 UTC 输出）。
/// 以前这里写死 +8（北京时间），对不在 +08:00 时区的用户是错的。
#[cfg(unix)]
fn local_utc_offset() -> i64 {
    #[repr(C)]
    struct Tm {
        tm_sec: i32,
        tm_min: i32,
        tm_hour: i32,
        tm_mday: i32,
        tm_mon: i32,
        tm_year: i32,
        tm_wday: i32,
        tm_yday: i32,
        tm_isdst: i32,
        tm_gmtoff: i64,
        tm_zone: *const i8,
    }
    extern "C" {
        fn localtime_r(t: *const i64, out: *mut Tm) -> *mut Tm;
    }
    let now = SystemTime::now().duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);
    let mut tm = Tm {
        tm_sec: 0, tm_min: 0, tm_hour: 0, tm_mday: 0, tm_mon: 0, tm_year: 0,
        tm_wday: 0, tm_yday: 0, tm_isdst: 0, tm_gmtoff: 0, tm_zone: core::ptr::null(),
    };
    let ok = unsafe { !localtime_r(&now, &mut tm).is_null() };
    if !ok {
        return 0;
    }
    let off = tm.tm_gmtoff;
    // 现实世界里没有超过 ±14 小时的时区，超了说明结构体布局对不上，退回 UTC
    if off.abs() > 14 * 3600 { 0 } else { off }
}

#[cfg(windows)]
fn local_utc_offset() -> i64 {
    #[repr(C)]
    #[derive(Default, Clone, Copy)]
    struct SystemTime {
        year: u16, month: u16, day_of_week: u16, day: u16,
        hour: u16, minute: u16, second: u16, milliseconds: u16,
    }
    #[link(name = "kernel32")]
    extern "system" {
        fn GetLocalTime(out: *mut SystemTime);
        fn GetSystemTime(out: *mut SystemTime);
    }
    // 用同一天同一时刻的本地时间与 UTC 时间相减得到偏移；跨月边界时换一天再算一次
    let as_minutes = |t: &SystemTime| -> i64 {
        (t.hour as i64) * 60 + (t.minute as i64)
    };
    let mut loc = SystemTime::default();
    let mut utc = SystemTime::default();
    unsafe {
        GetLocalTime(&mut loc);
        GetSystemTime(&mut utc);
    }
    if loc.day == utc.day {
        return (as_minutes(&loc) - as_minutes(&utc)) * 60;
    }
    // 跨日：本地比 UTC 落后还是超前，用日差补偿
    let day_delta = loc.day as i64 - utc.day as i64;
    (as_minutes(&loc) - as_minutes(&utc)) * 60 + day_delta * 86400
}

fn now_str() -> String {
    let secs = SystemTime::now().duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs() as i64).unwrap_or(0);
    let local = secs + local_utc_offset(); // 本地时区
    let days = local.div_euclid(86400);
    let rem = local.rem_euclid(86400);
    let (y, m, d) = civil_from_days(days);
    format!("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", y, m, d, rem / 3600, (rem % 3600) / 60, rem % 60)
}

fn parse_ymd(s: &str) -> Option<(i64, i64, i64)> {
    let t = s.trim();
    let t = t.split(' ').next().unwrap_or(t).trim();
    let norm: String = t.replace('年', "-").replace('月', "-").replace('日', "");
    let t = norm.as_str();
    let sep = if t.contains('-') {
        '-'
    } else if t.contains('/') {
        '/'
    } else if t.contains('.') {
        '.'
    } else {
        return None;
    };
    let parts: Vec<&str> = t.split(sep).collect();
    if parts.len() < 3 {
        return None;
    }
    let (y, m, d) = (
        parts[0].trim().parse::<i64>().ok()?,
        parts[1].trim().parse::<i64>().ok()?,
        parts[2].trim().parse::<i64>().ok()?,
    );
    if !(1..=12).contains(&m) || !(1..=31).contains(&d) {
        return None;
    }
    Some((y, m, d))
}

fn parse_hms(s: &str) -> Option<(i64, i64, i64)> {
    let t = s.trim();
    let t = t.rsplit(' ').next().unwrap_or(t).trim();
    if !t.contains(':') {
        return None;
    }
    let parts: Vec<&str> = t.split(':').collect();
    let g = |i: usize| -> Option<i64> { parts.get(i).and_then(|x| x.trim().parse::<i64>().ok()) };
    match parts.len() {
        3 => {
            let (h, m, s) = (g(0)?, g(1)?, g(2)?);
            if h > 23 || m > 59 || s > 59 { return None; }
            Some((h, m, s))
        }
        2 => {
            let (h, m) = (g(0)?, g(1)?);
            if h > 23 || m > 59 { return None; }
            Some((h, m, 0))
        }
        _ => None,
    }
}

fn fmt_hms(h: i64, m: i64, s: i64) -> String {
    format!("{:02}:{:02}:{:02}", h.rem_euclid(24), m.rem_euclid(60), s.rem_euclid(60))
}

/// 从文本解析出完整的时间分量；纯日期得到 0 时分秒，纯时间得到 0 年月日
fn parse_date_text(s: &str) -> Option<(i64, i64, i64, i64, i64, i64)> {
    let t = s.trim();
    if t.is_empty() || t.chars().count() > 32 {
        return None;
    }
    if t.chars().any(|c| c.is_ascii_alphabetic()) && !t.contains('年') {
        return None;
    }
    // 紧凑写法 20240305
    if t.len() == 8 && t.chars().all(|c| c.is_ascii_digit()) {
        let y: i64 = t[0..4].parse().ok()?;
        let m: i64 = t[4..6].parse().ok()?;
        let d: i64 = t[6..8].parse().ok()?;
        if (1900..=2099).contains(&y) && (1..=12).contains(&m) && (1..=31).contains(&d) {
            return Some((y, m, d, 0, 0, 0));
        }
        return None;
    }
    if let Some((a, b)) = t.split_once(' ') {
        if let Some((y, mo, d)) = parse_ymd(a) {
            let (h, mi, se) = parse_hms(b).unwrap_or((0, 0, 0));
            return Some((y, mo, d, h, mi, se));
        }
    }
    if let Some((y, mo, d)) = parse_ymd(t) {
        return Some((y, mo, d, 0, 0, 0));
    }
    if let Some((h, mi, se)) = parse_hms(t) {
        return Some((0, 0, 0, h, mi, se));
    }
    None
}

/// Excel 日期序列号 → 年月日时分秒
fn serial_to_parts(v: f64) -> (i64, i64, i64, i64, i64, i64) {
    let days = v.floor() as i64;
    let frac = v - days as f64;
    let (y, m, d) = civil_from_days(-25569 + days);
    let secs_total = (frac * 86400.0).round() as i64;
    (y, m, d, secs_total / 3600, (secs_total % 3600) / 60, secs_total % 60)
}

fn describe(c: &Cell) -> String {
    match c {
        Cell::Empty => "(空)".into(),
        Cell::Err(s) if s.is_empty() => "(Excel 错误值)".into(),
        other => format!("'{}'", truncate_disp(&other.text(), 40)),
    }
}

// ===========================================================================
// 行转换器
// ===========================================================================
struct Conv {
    bases: Vec<String>,
    specs: Vec<TypeSpec>,
    empty_as_null: String,
}

impl Conv {
    fn new(cols: &[Col], opts: &Opts) -> Self {
        let mut bases = Vec::new();
        let mut specs = Vec::new();
        for c in cols {
            let (b, s) = parse_type_spec(&c.mysql_type);
            bases.push(b);
            specs.push(s);
        }
        Conv { bases, specs, empty_as_null: opts.empty_as_null.clone() }
    }

    fn empty_literal(&self, base: &str) -> String {
        match self.empty_as_null.as_str() {
            "always" => "NULL".into(),
            "never" => if is_string_type(base) { "''".into() } else { "NULL".into() },
            _ => if is_string_type(base) { "''".into() } else { "NULL".into() },
        }
    }

    fn convert(&self, cells: &[Cell]) -> Result<(Vec<String>, String), String> {
        let mut lits = Vec::with_capacity(self.bases.len());
        let mut warns: Vec<String> = Vec::new();
        for i in 0..self.bases.len() {
            let cell = cells.get(i).unwrap_or(&EMPTY_CELL);
            let (lit, warn) = self.one(cell, &self.bases[i], self.specs[i])?;
            lits.push(lit);
            if let Some(w) = warn {
                warns.push(w);
            }
        }
        Ok((lits, warns.join(sep2())))
    }

    fn one(&self, c: &Cell, base: &str, spec: TypeSpec) -> Result<(String, Option<String>), String> {
        match c {
            Cell::Empty => return Ok((self.empty_literal(base), None)),
            Cell::Err(_) => return Ok((self.empty_literal(base), Some(t("Excel 错误值，已按空处理")))),
            _ => {}
        }

        // ---- 整数 ----
        if matches!(base, "tinyint" | "smallint" | "mediumint" | "int" | "integer" | "bigint") {
            if let Cell::Bool(b) = c {
                return Ok((if *b { "1" } else { "0" }.into(), None));
            }
            let n: Option<i128> = match c {
                Cell::Int(v) => Some(*v),
                Cell::Dec(f) | Cell::Sci(f) => {
                    if f.fract() != 0.0 {
                        return Err(format!("值为 {}，含小数，无法写入整数型字段", fmt_float(*f)));
                    }
                    Some(*f as i128)
                }
                Cell::Text(s) => s.trim().parse::<i128>().ok(),
                _ => None,
            };
            let n = match n {
                Some(v) => v,
                None => return Err(format!("值为 {}，不是整数，无法写入整数型字段", describe(c))),
            };
            if let TypeSpec::Int(lo, hi) = spec {
                if n < lo || n > hi {
                    return Err(format!("数值 {} 超出 {} 范围({}~{})", n, base, lo, hi));
                }
            }
            return Ok((n.to_string(), None));
        }

        // ---- 数值 ----
        if matches!(base, "decimal" | "numeric" | "double" | "float") {
            if let Cell::Bool(b) = c {
                return Ok((if *b { "1" } else { "0" }.into(), None));
            }
            let raw = match c {
                Cell::Int(v) => v.to_string(),
                Cell::Dec(f) | Cell::Sci(f) => fmt_float(*f),
                Cell::Text(s) => s.trim().to_string(),
                _ => return Err(format!("值为 {}，不是数值，无法写入 {} 字段", describe(c), base)),
            };
            let fv: f64 = raw.parse().map_err(|_| format!("值为 '{}'，不是合法数值", raw))?;
            if !fv.is_finite() {
                return Err(format!("数值 '{}' 不是有限数值", raw));
            }
            if base == "decimal" || base == "numeric" {
                let (p, sc) = match spec {
                    TypeSpec::Dec(p, s) => (p, s),
                    _ => (10, 0),
                };
                let body = raw.trim_start_matches(['+', '-']).to_string();
                let is_sci = body.contains('e') || body.contains('E');
                let (ip, fp) = if is_sci {
                    (format!("{:.0}", fv.abs().trunc()), String::new())
                } else {
                    match body.split_once('.') {
                        Some((a, b)) => (a.to_string(), b.to_string()),
                        None => (body.clone(), String::new()),
                    }
                };
                let ip = if ip.is_empty() { "0".to_string() } else { ip };
                if ip.len() as i64 > p - sc {
                    return Err(format!("数值 {} 整数位超过 decimal({},{}) 允许的 {} 位", raw, p, sc, p - sc));
                }
                if fp.len() as i64 > sc {
                    return Err(format!(
                        "数值 {} 小数位超过 decimal({},{}) 允许的 {} 位（可加大扫描样本后重试）",
                        raw, p, sc, sc));
                }
                return Ok((if is_sci { format!("{:.*}", sc as usize, fv) } else { raw }, None));
            }
            return Ok((format!("{}", fv), None));
        }

        // ---- 日期时间 ----
        if matches!(base, "date" | "datetime" | "time") {
            return self.as_temporal(c, base);
        }

        // ---- 字符串 ----
        let s = c.text();
        match base {
            "varchar" | "char" => {
                if let TypeSpec::Str(n) = spec {
                    let len = s.chars().count() as i64;
                    if len > n {
                        return Err(format!(
                            "内容长度 {} 超过 {}({}) 限制（样本未覆盖到，可加 --full-scan 重新扫描或先处理数据）",
                            len, base, n));
                    }
                }
            }
            "text" | "mediumtext" => {
                if s.len() > 65535 {
                    return Err(t("内容超过 text 上限(65535 字节)，建议改用 longtext"));
                }
            }
            _ => {}
        }
        Ok((sql_lit(&sql_safe(&s)), None))
    }

    fn as_temporal(&self, c: &Cell, base: &str) -> Result<(String, Option<String>), String> {
        let mut warn: Option<String> = None;
        let (y, mo, d, h, mi, se) = match c {
            Cell::Date(s) => {
                let (y, m, d) = parse_ymd(s).ok_or_else(|| format!("值为 '{}'，无法解析为 {}", s, base))?;
                (y, m, d, 0i64, 0i64, 0i64)
            }
            Cell::DateTime(s) => {
                let (y, m, d) = parse_ymd(s).ok_or_else(|| format!("值为 '{}'，无法解析为 {}", s, base))?;
                let (h, mi, se) = parse_hms(s).unwrap_or((0, 0, 0));
                (y, m, d, h, mi, se)
            }
            Cell::Time(s) => {
                let (h, mi, se) = parse_hms(s)
                    .ok_or_else(|| format!("值为 '{}'，无法解析为 {}", s, base))?;
                return Ok((sql_lit(&fmt_hms(h, mi, se)), None));
            }
            Cell::Text(s) => parse_date_text(s)
                .ok_or_else(|| format!("值为 '{}'，无法解析为 {}", truncate_disp(s, 40), base))?,
            Cell::Int(v) => {
                warn = Some(t("按 Excel 日期序列号转换"));
                serial_to_parts(*v as f64)
            }
            Cell::Dec(f) | Cell::Sci(f) => {
                warn = Some(t("按 Excel 日期序列号转换"));
                serial_to_parts(*f)
            }
            _ => return Err(format!("值为 {}，无法解析为 {}", describe(c), base)),
        };

        if base == "date" {
            if y <= 0 {
                return Err(format!("值为 {}，不含日期部分，无法写入 date", describe(c)));
            }
            if !matches!(c, Cell::Text(_)) && (h != 0 || mi != 0 || se != 0) {
                warn.get_or_insert_with(|| t("已丢弃时间部分"));
            }
            return Ok((sql_lit(&format!("{:04}-{:02}-{:02}", y, mo, d)), warn));
        }
        if base == "time" {
            return Ok((sql_lit(&fmt_hms(h, mi, se)), warn));
        }
        if y <= 1900 {
            return Err(format!("日期 {} 早于 1900 年，无法写入 datetime", y));
        }
        Ok((sql_lit(&format!("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", y, mo, d, h, mi, se)), warn))
    }
}


// ===========================================================================
// 抽样
// ===========================================================================
struct Plan {
    mode: &'static str,
    step: i64,
    offset: i64,
    ratio: f64,
    sampled: i64,
    /// 开头这一段行无条件纳入抽样（前 warm 行），避免前几行代表性不足
    warm: i64,
    /// 最多读取多少行（与 Python 版一致，大文件可提前收工）
    plan_rows: i64,
}

fn next_prime(n: i64) -> i64 {
    if n <= 2 {
        return 2;
    }
    let mut n = if n % 2 == 0 { n + 1 } else { n };
    loop {
        let mut d = 3;
        let mut is_prime = true;
        while d * d <= n {
            if n % d == 0 {
                is_prime = false;
                break;
            }
            d += 2;
        }
        if is_prime {
            return n;
        }
        n += 2;
    }
}

/// 抽样步长：取「最接近 1/ratio 的质数」。
///
/// 为什么不用精确的 1/ratio？为了避开与数据自身周期的共振。步长取 10 时，如果
/// 数据恰好以 50 行为周期，被抽到的行号模 50 永远只落在同样的 5 个余数上，会
/// 系统性漏掉一部分取值（例如金额列只在别的余数上出现两位小数），把类型判错，
/// 转换阶段大批行失败。质数步长与常见周期互质，能均匀覆盖。
fn sample_step(ratio: f64) -> i64 {
    let r = ratio.clamp(0.001, 1.0);
    next_prime(((1.0 / r).round() as i64).max(2))
}

fn choose_sampling(total: i64, opts: &Opts) -> Plan {
    if opts.full_scan || total <= opts.sample_threshold {
        return Plan { mode: "full", step: 1, offset: 0, ratio: 1.0, sampled: total,
                      warm: 0, plan_rows: total.max(0) };
    }
    let ratio = opts.sample_ratio.clamp(0.001, 1.0);
    let step = sample_step(ratio);
    let warm = opts.sample_min.max(0);
    let mut n = opts.sample_min.max((total as f64 * ratio) as i64);
    n = n.min(opts.sample_max).min(total);
    let plan_rows = total.min(warm.max(n) * step);
    // 起点：给了 --seed 就直接用它定位（seed % step），不同实现之间可对齐；
    // 没给就随机取，避免每次跑都只抽同一批行。
    let offset = match opts.seed {
        Some(seed) => (seed % step as u64) as i64,
        None => {
            let t = SystemTime::now().duration_since(UNIX_EPOCH)
                .map(|d| d.as_nanos() as u64).unwrap_or(0);
            (t % step as u64) as i64
        }
    };
    Plan { mode: "sample", step, offset, ratio, sampled: n, warm, plan_rows }
}

// ===========================================================================
// 探测
// ===========================================================================
struct Probe {
    kind: String,
    n_rows: i64,
    n_cols: i64,
    sheets: Vec<String>,
    merges: Vec<(i64, i64, i64, i64)>,
    formula: bool,
    hidden_cols: Vec<i64>,
    hidden_rows: Vec<i64>,
    embed_img: bool,
    /// 字段名行落在纵向合并区时，从合并区左上角取到的值：(列号(1起), 值)
    hfill: Vec<(usize, String)>,
}

fn do_probe(py: &str, opts: &Opts) -> Result<Probe, CliError> {
    let args = vec!["probe".to_string(), opts.path.clone(), opts.sheet.clone(),
                    opts.header_row.to_string()];
    let (code, stdout, stderr) = run_helper_once(py, &args, None)?;
    if let Some((code, msg)) = fatal_from_stdout(&stdout) {
        return Err(CliError::new(msg, code));
    }
    if code != 0 {
        return Err(CliError::general(format!(
            "读取 Excel 失败（Python 助手退出码 {}）\n{}", code, stderr_tail(&stderr, 6))));
    }
    let mut p = Probe {
        kind: String::new(), n_rows: 0, n_cols: 0,
        sheets: Vec::new(), merges: Vec::new(), formula: false,
        hidden_cols: Vec::new(), hidden_rows: Vec::new(), embed_img: false,
        hfill: Vec::new(),
    };
    for line in stdout.lines() {
        match parse_event(line) {
            Ev::Kind(k) => p.kind = k,
            Ev::SheetInfo(r, c) => {
                p.n_rows = r;
                p.n_cols = c;
            }
            Ev::Sheets(s) => p.sheets = s,
            Ev::Merges(m) => p.merges = m,
            Ev::Formula(f) => p.formula = f,
            Ev::Hidden(c, r) => {
                p.hidden_cols = c;
                p.hidden_rows = r;
            }
            Ev::EmbedImg(b) => p.embed_img = b,
            Ev::Hfill(v) => p.hfill = v,
            _ => {}
        }
    }
    if p.n_rows == 0 && p.n_cols == 0 {
        return Err(CliError::general(format!(
            "无法解析 Excel 结构（Python 助手无有效输出）\n{}", stderr_tail(&stderr, 6))));
    }
    Ok(p)
}

/// 输出目录不存在时给一条能看懂的错误，而不是把系统 errno 原样丢出来。
fn ensure_parent(path: &str) -> Result<(), CliError> {
    if let Some(d) = Path::new(path).parent() {
        if !d.as_os_str().is_empty() && !d.is_dir() {
            return Err(CliError::new(
                tf("输出目录不存在: {}（请先创建该目录，或用 --out / --err-file / --report 指定别处）",
                   &[&d.display().to_string()]),
                EXIT_USAGE));
        }
    }
    Ok(())
}

/// 原子输出守卫：先写 <目标>.part，成功提交时才改名到目标。
/// 中途失败（读表异常、助手崩溃、磁盘写满）会让守卫在析构时删掉半截文件 ——
/// 既不会留下残缺的 sql 让人误以为转换成功，也不会破坏上一次生成的成果。
/// 改名在 Unix 上是原子的；Windows 上先把旧目标删掉再改（失败时旧文件已不在，
/// 但至少不会出现两个半截文件）。
struct TmpOut {
    tmp: String,
    dest: String,
    armed: bool,
}

impl TmpOut {
    fn new(dest: &str) -> Self {
        TmpOut { tmp: format!("{}.part", dest), dest: dest.to_string(), armed: true }
    }
    fn tmp(&self) -> &str {
        &self.tmp
    }
    fn commit(mut self) -> io::Result<()> {
        if cfg!(windows) && Path::new(&self.dest).exists() {
            let _ = std::fs::remove_file(&self.dest);
        }
        match std::fs::rename(&self.tmp, &self.dest) {
            Ok(()) => {}
            Err(_) if !cfg!(windows) => {
                // Unix 上 rename 失败只可能是跨文件系统（--out 指到别的挂载点）。
                // 退化成复制：慢一点，但结果正确。
                std::fs::copy(&self.tmp, &self.dest)?;
                let _ = std::fs::remove_file(&self.tmp);
            }
            Err(e) => return Err(e),
        }
        self.armed = false;
        Ok(())
    }
}

impl Drop for TmpOut {
    fn drop(&mut self) {
        if self.armed {
            let _ = std::fs::remove_file(&self.tmp);
        }
    }
}

fn default_side_path(src: &str, name: &str) -> String {
    let p = Path::new(src);
    let dir = p.parent().filter(|d| !d.as_os_str().is_empty())
        .map(|d| d.to_path_buf())
        .unwrap_or_else(|| PathBuf::from("."));
    dir.join(name).to_string_lossy().into_owned()
}

// ===========================================================================
// 写 SQL
// ===========================================================================
fn write_sql_header(w: &mut impl Write, opts: &Opts, cols: &[Col]) -> io::Result<()> {
    let line = "─".repeat(60);
    let fname = Path::new(&opts.path).file_name()
        .map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
    writeln!(w, "-- {}", line)?;
    writeln!(w, "{}", tf("-- 由 xlsxtomysql {} 生成", &[&VERSION]))?;
    writeln!(w, "{}", tf("-- 源文件       : {}", &[&fname]))?;
    writeln!(w, "{}", tf("-- 工作表       : {}", &[&opts.sheet]))?;
    writeln!(w, "{}", tf("-- 字段名所在行 : {}", &[&opts.header_row]))?;
    writeln!(w, "{}", tf("-- 第一条数据行 : {}", &[&opts.data_row]))?;
    writeln!(w, "{}", tf("-- 字段数       : {}", &[&cols.len()]))?;
    writeln!(w, "{}", tf("-- 生成时间     : {}", &[&now_str()]))?;
    writeln!(w, "-- {}", line)?;
    writeln!(w)?;
    writeln!(w, "SET NAMES utf8mb4;")?;
    writeln!(w)?;
    if opts.drop_table {
        writeln!(w, "DROP TABLE IF EXISTS {};", ident(&opts.table))?;
        writeln!(w)?;
    }
    writeln!(w, "CREATE TABLE IF NOT EXISTS {} (", ident(&opts.table))?;
    let mut defs: Vec<String> = Vec::new();
    if let Some(id) = opts.add_id.as_ref() {
        defs.push(format!("  {} bigint NOT NULL AUTO_INCREMENT", ident(id)));
    }
    for c in cols {
        let mut d = format!("  {} {}", ident(&c.name), c.mysql_type);
        if opts.not_null {
            d.push_str(" NOT NULL");
        }
        let cmt = if !c.orig.is_empty() && c.orig != c.name {
            c.orig.clone()
        } else {
            String::new()
        };
        if !cmt.is_empty() {
            d.push_str(&format!(" COMMENT '{}'", comment_text(&cmt)));
        }
        defs.push(d);
    }
    let pk = opts.add_id.as_ref().or(opts.primary_key.as_ref());
    if let Some(pk) = pk {
        defs.push(format!("  PRIMARY KEY ({})", ident(pk)));
    }
    writeln!(w, "{}", defs.join(",\n"))?;
    let mut tail = String::from(") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci");
    if let Some(tc) = opts.table_comment.as_ref() {
        tail.push_str(&format!(" COMMENT='{}'", comment_text(tc)));
    }
    writeln!(w, "{};", tail)?;
    writeln!(w)?;
    Ok(())
}

fn write_insert_batch(
    w: &mut impl Write, table: &str, cols: &[Col],
    rows: &[Vec<String>], ignore: bool,
) -> io::Result<()> {
    let names: Vec<String> = cols.iter().map(|c| ident(&c.name)).collect();
    writeln!(
        w, "{} INTO {} ({}) VALUES",
        if ignore { "INSERT IGNORE" } else { "INSERT" },
        ident(table), names.join(",")
    )?;
    let last = rows.len().saturating_sub(1);
    for (i, r) in rows.iter().enumerate() {
        writeln!(w, "  ({}){}", r.join(","), if i == last { ";" } else { "," })?;
    }
    Ok(())
}

fn append_tail_comment(path: &str, ok: i64, failed: i64, elapsed: f64) -> io::Result<()> {
    use std::fs::OpenOptions;
    let mut f = OpenOptions::new().append(true).open(path)?;
    writeln!(f)?;
    writeln!(f, "-- {}", "─".repeat(60))?;
    writeln!(f, "{}", tf("-- 完成: 共 {} 行, 成功 {} 行, 失败 {} 行, 用时 {}",
                         &[&human_num(ok + failed), &human_num(ok),
                           &human_num(failed), &fmt_dur(elapsed)]))?;
    writeln!(f, "-- {}", "─".repeat(60))?;
    Ok(())
}

// ===========================================================================
// 失败行导出
// ===========================================================================
fn write_errrows(py: &str, path: &str, cols: &[Col], fails: &[(i64, Vec<String>, String)])
    -> Result<bool, CliError>
{
    if fails.is_empty() {
        return Ok(false);
    }
    let mut payload = String::new();
    payload.push('H');
    payload.push('\t');
    payload.push_str(&t("原行号"));
    for c in cols {
        payload.push('\t');
        payload.push_str(&escape(if c.orig.is_empty() { &c.name } else { &c.orig }));
    }
    payload.push('\t');
    payload.push_str(&t("失败原因"));
    payload.push('\n');
    for (row, texts, reason) in fails {
        payload.push('R');
        payload.push('\t');
        payload.push_str(&row.to_string());
        for t in texts {
            payload.push('\t');
            payload.push_str(&escape(t));
        }
        payload.push('\t');
        payload.push_str(&escape(reason));
        payload.push('\n');
    }
    let args = vec!["errrows".to_string(), path.to_string()];
    let (code, stdout, stderr) = run_helper_once(py, &args, Some(&payload))?;
    if let Some((_code, m)) = fatal_from_stdout(&stdout) {
        return Err(CliError::general(format!("导出失败行失败: {}", m)));
    }
    if code != 0 {
        return Err(CliError::general(format!(
            "导出失败行失败（退出码 {}）\n{}", code, stderr_tail(&stderr, 5))));
    }
    Ok(true)
}

// ===========================================================================
// 报表
// ===========================================================================
fn print_issues(out: &mut Out, st: Style, issues: &[Issue]) {
    let blockers: Vec<&Issue> = issues.iter().filter(|i| i.level == "block").collect();
    let warns: Vec<&Issue> = issues.iter().filter(|i| i.level == "warn").collect();
    if blockers.is_empty() && warns.is_empty() {
        return;
    }
    out.blank();
    if !blockers.is_empty() {
        out.line(&st.red(&format!(
            "发现 {} 处非标准格式（会造成数据错位，需处理源文件后重跑）：", blockers.len())));
        let rows: Vec<Vec<String>> = blockers.iter()
            .map(|i| vec![i.where_(), i.title.clone(), i.detail.clone()]).collect();
        out.line(&render_table(&["位置", "问题", "说明"], &rows));
    }
    if !warns.is_empty() {
        out.blank();
        out.line(&st.yellow(&format!("另有 {} 处提醒（不阻断转换，建议检查）：", warns.len())));
        let rows: Vec<Vec<String>> = warns.iter()
            .map(|i| vec![i.where_(), i.title.clone(), i.detail.clone()]).collect();
        out.line(&render_table(&["位置", "提醒", "说明"], &rows));
    }
}

fn report_types(out: &mut Out, st: Style, cols: &[Col]) {
    out.blank();
    out.line(&st.bold("结果表 2 · 新字段名称及类型"));
    let rows: Vec<Vec<String>> = cols.iter().map(|c| vec![
        c.name.clone(),
        c.mysql_type.clone(),
        c.reason.clone(),
        if c.note.is_empty() { "-".to_string() } else { c.note.clone() },
    ]).collect();
    out.line(&render_table(&["新字段名", "MySQL 类型", "推断依据", "字段名处理"], &rows));
}

fn normalize_reason(r: &str) -> String {
    let mut out = String::new();
    let mut in_quote = false;
    let mut prev_digit = false;
    for ch in r.chars() {
        if in_quote {
            if ch == '」' {
                out.push_str("」");
                in_quote = false;
            }
            continue;
        }
        match ch {
            '「' => {
                out.push_str("「字段」");
                in_quote = true;
                prev_digit = false;
            }
            c if c.is_ascii_digit() => {
                if !prev_digit {
                    out.push('N');
                }
                prev_digit = true;
            }
            _ => {
                prev_digit = false;
                out.push(ch);
            }
        }
    }
    out
}

struct RunInfo<'a> {
    opts: &'a Opts,
    kind: &'a str,
    out_sql: &'a str,
    total_rows: i64,
    ok_rows: i64,
    warns: i64,
    skipped: i64,
    scanned: i64,
    plan: &'a Plan,
    elapsed: f64,
}

fn report_result(out: &mut Out, st: Style, info: &RunInfo, fails: &[(i64, Vec<String>, String)], err_file: Option<&str>) {
    let opts = info.opts;
    out.blank();
    out.line(&st.bold("结果表 1 · 转换结果"));
    let scan_desc = if info.plan.mode == "full" {
        format!("全量扫描 {} 行", human_num(info.scanned))
    } else {
        format!("抽样跳跃扫描（每 {} 行取 1 行，判定 {} 行）", info.plan.step, human_num(info.scanned))
    };
    let mut rows: Vec<Vec<String>> = vec![
        vec!["源文件".into(), Path::new(&opts.path).file_name().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default()],
        vec!["文件格式".into(), info.kind.to_string()],
        vec!["工作表".into(), opts.sheet.clone()],
        vec!["字段名所在行".into(), opts.header_row.to_string()],
        vec!["第一条数据所在行".into(), opts.data_row.to_string()],
        vec!["共多少行数据".into(), human_num(info.total_rows)],
        vec!["成功行数".into(), human_num(info.ok_rows)],
        vec!["失败行数".into(), human_num(fails.len() as i64)],
        vec!["警告行数".into(), human_num(info.warns)],
        vec!["跳过空行".into(), human_num(info.skipped)],
        vec!["类型扫描方式".into(), scan_desc],
        vec!["耗时".into(), fmt_dur(info.elapsed)],
        vec!["输出文件".into(), info.out_sql.to_string()],
    ];
    if let Some(f) = err_file {
        rows.push(vec!["失败行文件".into(), f.to_string()]);
    }
    out.line(&render_table(&["项目", "值"], &rows));

    if !fails.is_empty() {
        out.blank();
        out.line(&st.yellow(&format!(
            "失败明细（原文件行号，最多显示 {} 条）", opts.print_errors)));
        let shown: Vec<Vec<String>> = fails.iter().take(opts.print_errors)
            .map(|(r, _, e)| vec![human_num(*r), e.clone()]).collect();
        out.line(&render_table(&["原行号", "失败原因"], &shown));
        if fails.len() > opts.print_errors {
            out.line(&st.dim(&format!("  ... 其余 {} 行见 {}", fails.len() - opts.print_errors, "errrows.xlsx")));
        }
        let mut agg: HashMap<String, i64> = HashMap::new();
        for (_, _, e) in fails {
            *agg.entry(normalize_reason(e)).or_insert(0) += 1;
        }
        let mut list: Vec<(String, i64)> = agg.into_iter().collect();
        list.sort_by(|a, b| b.1.cmp(&a.1));
        out.blank();
        let rows: Vec<Vec<String>> = list.iter().map(|(k, v)| vec![k.clone(), human_num(*v)]).collect();
        out.line(&render_table(&["失败原因归类", "行数"], &rows));
    }
}

// ===========================================================================
// 入口
// ===========================================================================
fn main() {
    // 顺序有讲究：先恢复 SIGPIPE，再接管控制台编码，最后才可能开始输出
    restore_sigpipe();
    console_setup();
    // args() 遇到非法 UTF-8 会 panic（Linux 文件名是任意字节），这里做有损转换，
    // 宁可个别字节显示成 ? 也不要直接崩。
    let argv: Vec<String> = std::env::args_os()
        .skip(1)
        .map(|a| a.to_string_lossy().into_owned())
        .collect();
    // 语言必须先于任何输出定下来：--lang=xx > XLSXTOMYSQL_LANG > LC_ALL/LC_MESSAGES/LANG
    let mut lang = String::new();
    let mut li = 0usize;
    while li < argv.len() {
        if let Some(v) = argv[li].strip_prefix("--lang=") {
            lang = v.to_string();
            break;
        }
        if argv[li] == "--lang" {
            lang = argv.get(li + 1).cloned().unwrap_or_default();
            break;
        }
        li += 1;
    }
    set_lang(&lang);
    let opts = match parse_args(argv) {
        Ok(o) => o,
        Err(e) => {
            eprintln!("{} {}", t("错误:"), e.msg);
            std::process::exit(e.code);
        }
    };
    let st = Style { color: opts.color && io::stdout().is_terminal() };
    let mut out = Out::new(opts.report.is_some());
    let code = match run(&opts, st, &mut out) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("{} {}", st.red("错误:"), e.msg);
            e.code
        }
    };
    if let Some(p) = opts.report.as_ref() {
        match ensure_parent(p) {
            Err(e) => eprintln!("{} {}", st.red("错误:"), e.msg),
            Ok(()) => match out.save(p) {
                Ok(_) => eprintln!("{}", tf("报表已保存: {}", &[&p])),
                Err(e) => eprintln!("{}", tf("报表写入失败: {} ({})", &[&p, &e])),
            },
        }
    }
    std::process::exit(code);
}

fn run(opts: &Opts, st: Style, out: &mut Out) -> Result<i32, CliError> {
    let py = find_python(opts);
    if !Path::new(&opts.path).exists() {
        return Err(CliError::new(format!("文件不存在: {}", opts.path), EXIT_USAGE));
    }
    out.line(&st.bold(&format!("xlsxtomysql {} —— Excel → MySQL", VERSION)));
    out.blank();

    // ---------------- 阶段 0: 探测 ----------------
    let probe = do_probe(&py, opts)?;
    out.line(&format!("{} {}", st.cyan("文件格式:"), probe.kind));
    out.line(&format!("{} {}  （{} 行 × {} 列）", st.cyan("工作表  :"),
                      opts.sheet, human_num(probe.n_rows), probe.n_cols));
    if !probe.sheets.is_empty() && probe.sheets.len() > 1 {
        out.line(&st.dim(&tf("  可用工作表: {}", &[&probe.sheets.join(sep())])));
    }
    if probe.formula {
        out.line(&st.yellow("  ⚠ 含公式；当前按「计算结果」读取，若整列为空请先用 Excel 打开保存一次"));
    }
    if !probe.merges.is_empty() {
        out.line(&format!("  检测到 {} 处合并单元格", probe.merges.len()));
        if opts.fill_down {
            out.line(&st.dim("  数据区的合并单元格会按左上角内容自动补齐整块（--no-fill-down 可关闭）"));
        }
    }
    if probe.embed_img {
        out.line(&st.yellow("  ⚠ 检测到「置于单元格内」的图片（WPS 的 DISPIMG / Excel 单元格图片）；\
图片本身无法写入 SQL，对应单元格会变成空值"));
    }
    if !probe.hidden_cols.is_empty() {
        let shown: Vec<String> = probe.hidden_cols.iter().take(8)
            .map(|c| col_name(*c as usize)).collect();
        out.line(&st.yellow(&format!(
            "  ⚠ 检测到 {} 个隐藏列（{}{}），隐藏只是不显示，内容仍会被导出；如需排除请先删除该列",
            probe.hidden_cols.len(), shown.join(sep()),
            if probe.hidden_cols.len() > 8 { "…" } else { "" })));
    }
    if !probe.hidden_rows.is_empty() {
        out.line(&st.yellow(&format!(
            "  ⚠ 检测到 {} 个隐藏行，其内容同样会被导出", probe.hidden_rows.len())));
    }

    // ---------------- 行范围 ----------------
    let mut last_row = probe.n_rows;
    if let Some(want_rows) = opts.rows {
        let want = opts.data_row + want_rows - 1;
        if last_row > 0 && want > last_row {
            out.line(&st.yellow(&format!(
                "  ⚠ 「共几行」={} 超出工作表范围，已自动截断到第 {} 行",
                human_num(want_rows), human_num(last_row))));
            last_row = last_row.min(want);
        } else if last_row == 0 {
            last_row = want;
        } else {
            last_row = want;
        }
    }
    if last_row < opts.data_row {
        return Err(CliError::new(format!(
            "数据区为空: 起始行 {} 已超出工作表范围（共 {} 行）",
            opts.data_row, human_num(probe.n_rows)), EXIT_USAGE));
    }
    let total_rows = last_row - opts.data_row + 1;

    // ---------------- 阶段 1: 扫描 ----------------
    let plan = choose_sampling(total_rows, opts);
    let desc = if plan.mode == "full" {
        tf("全量扫描 {} 行", &[&human_num(total_rows)])
    } else {
        let pct = format!("{:.0}", plan.ratio * 100.0);
        tf("抽样扫描: 共 {} 行，按 {}% 跳跃抽样，每 {} 行取 1 行，实际判定约 {} 行",
           &[&human_num(total_rows), &pct, &plan.step, &human_num(plan.sampled)])
    };
    out.blank();
    out.line(&st.bold("【1/2】预扫描 —— 识别字段类型 & 检查表格格式"));
    out.line(&format!("{} {}", st.cyan("扫描方式:"), desc));

    let n_cols = probe.n_cols.max(1);
    // 助手的 max_read 是「从字段名行起一共读多少行」，而 plan_rows 是「数据行数」，
    // 两者差一个表头到数据区的行距；全量扫描时干脆不设上限，避免漏掉最后一行。
    let max_read = if plan.mode == "full" {
        0
    } else {
        plan.plan_rows + (opts.data_row - opts.header_row)
    };
    let dump_args = vec![
        "dump".to_string(), opts.path.clone(), opts.sheet.clone(),
        opts.header_row.to_string(), opts.data_row.to_string(), last_row.to_string(),
        plan.step.to_string(), plan.offset.to_string(),
        if opts.hints { "1".into() } else { "0".into() },
        n_cols.to_string(),
        plan.warm.to_string(), max_read.to_string(),
    ];
    let label_scan = t("扫描中");
    let mut s1 = Stream::start(&py, &dump_args, &label_scan, opts.progress)?;

    let mut header: Option<Vec<Cell>> = None;
    let mut cols: Vec<Col> = Vec::new();
    let mut accs: Vec<ColAcc> = Vec::new();
    let mut checker: Option<Checker> = None;
    let mut scanned: i64 = 0;
    let mut helper_fatal: Option<(i32, String)> = None;
    let mut pct_cells: i64 = 0;
    let mut tz_dropped: Vec<(usize, i64)> = Vec::new();
    let mut scan_filler = if opts.fill_down {
        Some(MergeFiller::new(&probe.merges, opts.data_row))
    } else {
        None
    };

    loop {
        let ev = match s1.rx.recv() {
            Ok(e) => e,
            Err(_) => break,
        };
        match ev {
            Ev::Header(cells) => {
                let mut h = cells;
                let (touched, vertical) = fill_merged(&probe.merges, opts.header_row, &mut h,
                                                      &probe.hfill);
                // 与 Python 版一致：逐条说明"字段名取自哪个纵向合并区"
                for (r1, c1, r2, c2) in probe.merges.iter() {
                    if !(*r1 < opts.header_row && opts.header_row <= *r2) {
                        continue;
                    }
                    if !probe.hfill.iter().any(|(c, _)| *c as i64 == *c1) {
                        continue;
                    }
                    out.line(&format!(
                        "  {} 字段名行位于纵向合并区 {}{}:{}{} 内，已取合并区左上角的内容作为字段名",
                        st.cyan("▸"), col_name(*c1 as usize), r1,
                        col_name(*c2 as usize), r2));
                }
                if touched > 0 {
                    out.line(&format!("  {} 字段名行有 {} 处合并单元格，已自动填充{}",
                                      st.cyan("▸"), touched,
                                      if vertical > 0 {
                                          format!("（其中 {} 处为纵向合并）", vertical)
                                      } else {
                                          String::new()
                                      }));
                }
                let width = last_non_empty(&h);
                if width == 0 {
                    s1.kill();
                    return Err(CliError::new(format!(
                        "第 {} 行（字段名称所在行）没有任何内容，请检查行号参数是否正确",
                        opts.header_row), EXIT_USAGE));
                }
                let head = h[..width.min(h.len())].to_vec();
                cols = build_columns(&head, opts);
                accs = (0..cols.len()).map(|_| ColAcc::default()).collect();
                checker = Some(Checker::new(&head, width));
                header = Some(h);
            }
            Ev::Row(n, mut cells) => {
                if let Some(f) = scan_filler.as_mut() {
                    f.apply(n, &mut cells);
                }
                if let Some(ck) = checker.as_mut() {
                    if ck.add(n, &cells) {
                        scanned += 1;
                        for i in 0..accs.len() {
                            accs[i].add(cells.get(i).unwrap_or(&EMPTY_CELL));
                        }
                    }
                }
            }
            Ev::Pct(v) => pct_cells = v,
            Ev::Tz(v) => tz_dropped = v,
            Ev::Fatal(code, m) => helper_fatal = Some((code, m)),
            _ => {}
        }
    }
    let code1 = s1.child.wait().map(|s| s.code().unwrap_or(-1)).unwrap_or(-1);
    if let Some((code, m)) = helper_fatal {
        return Err(CliError::new(m, code));
    }
    if code1 != 0 {
        return Err(CliError::general(format!("读取阶段异常退出（退出码 {}）", code1)));
    }
    if header.is_none() {
        return Err(CliError::new(
            "未能读到字段名行，请检查「字段名称所在行」参数是否正确", EXIT_USAGE));
    }

    // 主键 / 自增列校验
    let names: HashSet<&str> = cols.iter().map(|c| c.name.as_str()).collect();
    if let Some(id) = opts.add_id.as_ref() {
        if names.contains(id.as_str()) {
            return Err(CliError::new(format!(
                "--add-id {} 与已有字段重名，请换一个列名（如 --add-id id）", id), EXIT_USAGE));
        }
    }
    if let Some(pk) = opts.primary_key.as_ref() {
        if !names.contains(pk.as_str()) {
            return Err(CliError::new(format!(
                "--primary-key {} 在字段名行中不存在。可用字段: {}",
                pk, cols.iter().map(|c| c.name.clone()).collect::<Vec<_>>().join(sep())), EXIT_USAGE));
        }
    }

    // 类型推断
    for i in 0..cols.len() {
        let (t, r) = resolve_type(&accs[i], &cols[i].name, opts);
        cols[i].mysql_type = t;
        cols[i].reason = r;
    }
    let mut ck = checker.expect("已读过表头");
    ck.finalize(last_row, opts.data_row);
    let mut issues = std::mem::take(&mut ck.issues);

    // 容易踩坑但程序无法自动修正的情况，一律提示出来
    if pct_cells > 0 {
        out.line(&st.yellow(&format!(
            "  ⚠ 检测到 {} 个百分比格式的单元格：Excel 里显示 12.5%，\
实际存储值是原始小数 0.125，入库的也是 0.125", human_num(pct_cells))));
    }
    let huge: Vec<String> = (0..cols.len()).filter(|i| accs[*i].huge_numbers > 0)
        .map(|i| format!("{}（{} 个）", cols[i].name, human_num(accs[i].huge_numbers)))
        .collect();
    if !huge.is_empty() {
        out.line(&st.yellow(&format!(
            "  ⚠ 以下字段有 15 位以上的数字，Excel 只能精确保存 15 位，末位可能已被改写为 0，\
请核对原始数据：{}", huge.join(sep()))));
    }
    let oor: Vec<String> = (0..cols.len()).filter(|i| accs[*i].out_of_range > 0)
        .map(|i| format!("{}（{} 个）", cols[i].name, human_num(accs[i].out_of_range)))
        .collect();
    if !oor.is_empty() {
        out.line(&st.yellow(&format!(
            "  ⚠ 以下字段有超出 MySQL 日期/时间范围的值，已整列按文本保存：{}", oor.join(sep()))));
    }
    let tz: Vec<String> = tz_dropped.iter()
        .filter(|(c, _)| *c >= 1 && *c <= cols.len())
        .map(|(c, n)| format!("{}（{} 个）", cols[*c - 1].name, human_num(*n)))
        .collect();
    if !tz.is_empty() {
        out.line(&st.yellow(&format!(
            "  ⚠ 以下字段的原值带时区偏移（如 2024-01-01T08:30:00+08:00），\
MySQL 的 datetime 不存时区，偏移量已被丢弃、只保留字面时间；\
如需按时区换算请先在 Excel 里统一：{}", tz.join(sep()))));
    }

    // 表头空缺提示
    for (i, c) in cols.iter().enumerate() {
        if c.orig.is_empty() {
            issues.push(Issue {
                level: "warn", row: opts.header_row,
                col: Some(i + 1),
                title: "字段名为空".into(),
                detail: format!("已自动命名为 `{}`（如需自定义请补全 {}{} 单元格）",
                                c.name, col_name(i + 1), opts.header_row),
            });
        }
    }

    print_issues(out, st, &issues);

    if opts.scan_only {
        report_types(out, st, &cols);
        out.blank();
        out.line(&st.cyan("仅扫描模式，未生成 sql。"));
        return Ok(EXIT_OK);
    }

    let blockers = issues.iter().filter(|i| i.level == "block").count();
    if blockers > 0 && !opts.force {
        out.blank();
        out.line(&st.red(&format!("发现 {} 处非标准格式，已中止转换（未生成 sql）。", blockers)));
        out.line("请按上面的位置提示处理源文件后重新执行；确需强行转换可追加 --force");
        return Ok(EXIT_FORMAT);
    }

    // ---------------- 阶段 2: 转换 ----------------
    let out_sql = opts.out.clone()
        .unwrap_or_else(|| default_side_path(&opts.path, &format!("{}.sql", opts.table)));
    let out_err = opts.err_file.clone()
        .unwrap_or_else(|| default_side_path(&opts.path, "errrows.xlsx"));

    ensure_parent(&out_sql)?;
    ensure_parent(&out_err)?;

    out.blank();
    out.line(&st.bold("【2/2】转换 —— 生成 SQL"));
    let conv = Conv::new(&cols, opts);
    let t0 = Instant::now();
    let guard = TmpOut::new(&out_sql);
    let file = File::create(guard.tmp())
        .map_err(|e| CliError::general(format!("无法写入 {}: {}", out_sql, e)))?;
    let mut w = BufWriter::with_capacity(1 << 20, file);
    write_sql_header(&mut w, opts, &cols)
        .map_err(|e| CliError::general(format!("写入 SQL 失败: {}", e)))?;

    let conv_args = vec![
        "dump".to_string(), opts.path.clone(), opts.sheet.clone(),
        opts.header_row.to_string(), opts.data_row.to_string(), last_row.to_string(),
        "1".to_string(), "0".to_string(),
        if opts.hints { "1".into() } else { "0".into() },
        n_cols.to_string(),
        "0".to_string(), "0".to_string(),
    ];
    let label_conv = t("转换中");
    let mut s2 = Stream::start(&py, &conv_args, &label_conv, opts.progress)?;

    let mut ok_rows = 0i64;
    let mut skipped_empty = 0i64;
    let mut warns_count = 0i64;
    let mut batch: Vec<Vec<String>> = Vec::new();
    let mut fails: Vec<(i64, Vec<String>, String)> = Vec::new();
    let mut io_err: Option<io::Error> = None;
    let mut helper_fatal2: Option<(i32, String)> = None;
    let mut conv_filler = if opts.fill_down {
        Some(MergeFiller::new(&probe.merges, opts.data_row))
    } else {
        None
    };

    loop {
        let ev = match s2.rx.recv() {
            Ok(e) => e,
            Err(_) => break,
        };
        if let Ev::Row(n, mut cells) = ev {
            if let Some(f) = conv_filler.as_mut() {
                f.apply(n, &mut cells);
            }
            if cells.iter().all(|c| c.is_empty()) {
                skipped_empty += 1;
                continue;
            }
            match conv.convert(&cells) {
                Ok((lits, warn)) => {
                    if !warn.is_empty() {
                        warns_count += 1;
                    }
                    batch.push(lits);
                    ok_rows += 1;
                    if batch.len() as i64 >= opts.batch_size {
                        if let Err(e) = write_insert_batch(&mut w, &opts.table, &cols, &batch, opts.insert_ignore) {
                            io_err = Some(e);
                            break;
                        }
                        batch.clear();
                    }
                }
                Err(reason) => {
                    let texts: Vec<String> = (0..cols.len())
                        .map(|i| cells.get(i).map(|c| c.text()).unwrap_or_default())
                        .collect();
                    fails.push((n, texts, reason));
                }
            }
        } else if let Ev::Fatal(code, m) = ev {
            helper_fatal2 = Some((code, m));
        }
    }

    if io_err.is_none() && !batch.is_empty() {
        if let Err(e) = write_insert_batch(&mut w, &opts.table, &cols, &batch, opts.insert_ignore) {
            io_err = Some(e);
        }
    }
    if io_err.is_some() {
        s2.kill();
    }
    let code2 = s2.child.wait().map(|s| s.code().unwrap_or(-1)).unwrap_or(-1);
    if let Some(e) = io_err {
        return Err(CliError::general(format!("写入 SQL 失败: {}", e)));
    }
    if let Some((code, m)) = helper_fatal2 {
        return Err(CliError::new(m, code));
    }
    if code2 != 0 {
        return Err(CliError::general(format!("读取阶段异常退出（退出码 {}）", code2)));
    }
    w.flush().map_err(|e| CliError::general(format!("写入 SQL 失败: {}", e)))?;
    drop(w);
    let merge_filled = conv_filler.as_ref().map(|f| f.filled).unwrap_or(0);
    if merge_filled > 0 {
        out.line(&format!("  {} 数据区合并单元格已补齐 {} 个空单元格（--no-fill-down 可关闭）",
                          st.cyan("▸"), human_num(merge_filled)));
    }
    let elapsed = t0.elapsed().as_secs_f64();
    // 尾部统计先写进临时文件，改名之后文件名才对得上
    append_tail_comment(guard.tmp(), ok_rows, fails.len() as i64, elapsed)
        .map_err(|e| CliError::general(format!("写入 SQL 尾部失败: {}", e)))?;
    guard.commit()
        .map_err(|e| CliError::general(format!("生成 {} 失败: {}", out_sql, e)))?;

    let wrote_err = write_errrows(&py, &out_err, &cols, &fails)?;

    let info = RunInfo {
        opts, kind: &probe.kind, out_sql: &out_sql,
        total_rows, ok_rows, warns: warns_count, skipped: skipped_empty,
        scanned, plan: &plan, elapsed,
    };
    report_result(out, st, &info, &fails, if wrote_err { Some(out_err.as_str()) } else { None });
    report_types(out, st, &cols);
    out.blank();
    out.line(&st.bold(&format!("完成: {}", out_sql)));
    Ok(EXIT_OK)
}
