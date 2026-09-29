# xlsxtomysql

**Excel → MySQL, from the command line. Turn `.xlsx` / `.xls` spreadsheets into `CREATE TABLE` + `INSERT` statements in one command.**

[![Release](https://img.shields.io/github/v/release/Paul-sinbud2004/xlsxtomysql?sort=semver)](https://github.com/Paul-sinbud2004/xlsxtomysql/releases)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![AUR](https://img.shields.io/aur/version/xlsxtomysql)](https://aur.archlinux.org/packages/xlsxtomysql)
[![Platform](https://img.shields.io/badge/platform-linux%20%7C%20macOS%20%7C%20windows-lightgrey.svg)](#cross-platform)
[![Zero crates](https://img.shields.io/badge/dependencies-none-success.svg)](#how-it-is-built)

A **single-file Rust program** — no crates, no Cargo dependency tree, one `rustc` command
produces one self-contained binary. It reads spreadsheet files through a small Python
helper embedded *inside the Rust source*, so nothing is written to disk at runtime.

[English](README.md) · [简体中文](README.zh-CN.md)

---

## Contents

- [Why](#why)
- [Features](#features)
- [Install](#install)
- [Quick start](#quick-start)
- [Command-line options](#command-line-options)
- [Messy spreadsheets](#messy-spreadsheets)
- [Type inference](#type-inference)
- [Large files and sampling](#large-files-and-sampling)
- [Language](#language)
- [Cross-platform](#cross-platform)
- [Exit codes](#exit-codes)
- [How it is built](#how-it-is-built)
- [FAQ](#faq)
- [Keywords](#keywords)

## Why

Getting a spreadsheet into MySQL is usually one of these: a GUI import wizard that
guesses wrong, a hand-written `CREATE TABLE` with the wrong types, or a throwaway script.
`xlsxtomysql` does the boring part properly:

- **Infers the schema** from the actual data, not from the first row — so a column that
  starts with `00123` does not become an integer that loses its leading zeros.
- **Refuses to silently mangle a messy sheet.** Blank rows in the middle, decorative
  rules, repeated headers and total rows are reported with their **original row numbers**,
  and the tool stops instead of producing a half-wrong table.
- **Scales to large files** with a sampled jump scan, so a 30,000-row sheet does not need
  a full pass before you can see the inferred types.
- **Exports the rows it could not convert** to `errrows.xlsx`, so nothing is quietly dropped.
- **Works the same on Linux, macOS and Windows**, in Chinese or English, automatically.

## Features

| | |
|---|---|
| **Input** | Excel 2007+ (`.xlsx`) and Excel 2003 (`.xls`); multiple worksheets; merged header cells; hidden rows/columns |
| **Output** | `CREATE TABLE` + batched `INSERT` statements, always `utf8mb4`; backticked identifiers; optional `DROP TABLE`, `INSERT IGNORE`, auto-increment primary key, table comment, `NOT NULL` |
| **Schema inference** | Per-column statistics (length distribution, integer width, decimal places, date/time values) combined with column-name semantics (`phone`, `id_card`, `amount`, `date`, …) |
| **Validation** | Blank rows inside the data region, separator/decoration rows, repeated headers, subtotal/total rows, data beyond the header column range — each reported by original row (and column) number |
| **Large files** | Automatic sampled jump scan above a row threshold, with a prime step size to avoid resonance with periodic data; reproducible seeds |
| **Failure handling** | Per-row failure reasons, grouped by cause in the report, plus an `errrows.xlsx` export |
| **Reports** | Console report (totals, per-row failures, inferred type ↔ reason table) that can be saved to a text file with `--report` |
| **Interface** | Chinese/English, auto-detected; `--help` cheat sheet and a long-form `--man` manual in both languages |
| **Packaging** | Single binary, one `rustc` command, no crates; man page; install scripts for Unix and Windows |

## Install

### Arch Linux (AUR)

```bash
yay -S xlsxtomysql      # or: paru -S xlsxtomysql
```

### Prebuilt binary

Download the tarball for your platform from
[Releases](https://github.com/Paul-sinbud2004/xlsxtomysql/releases) — it contains the
binary, both man pages, the docs and the install scripts:

```bash
tar xzf xlsxtomysql-<version>-<platform>.tar.gz
cd xlsxtomysql-<version>-<platform>
bash install.sh                       # install into ~/.local
bash install.sh --prefix /usr/local   # install system-wide (may need sudo)
bash install.sh --no-python           # skip creating a dedicated venv
```

On Windows, in the extracted folder:

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
powershell -ExecutionPolicy Bypass -File install.ps1 -Prefix "D:\tools"
```

Uninstalling is just deleting what was installed — the scripts touch nothing else:

```
<prefix>/bin/xlsxtomysql
<prefix>/share/man/man1/xlsxtomysql*.1
<prefix>/share/xlsxtomysql/
```

### From source

Only `rustc` is needed — **no Cargo, no crates, no network**:

```bash
git clone https://github.com/Paul-sinbud2004/xlsxtomysql.git
cd xlsxtomysql
bash build.sh --release          # opt-level=3 + LTO + stripped symbols
# or the single underlying command:
rustc --edition 2021 -O xlsxtomysql.rs -o xlsxtomysql
```

`cargo build --release` works too (`Cargo.toml` points the bin target at `xlsxtomysql.rs`).

### Runtime requirement

The Rust binary itself has **no** dependencies, but reading spreadsheets needs a Python
interpreter with:

| Package | Needed for | Required |
|---|---|---|
| `openpyxl` | reading `.xlsx` | yes |
| `xlrd` | reading Excel 2003 `.xls` | only for `.xls` |

Writing `errrows.xlsx` needs no third-party library — the helper assembles the OOXML with
the standard-library `zipfile`, creating and deleting no temporary files, so it also works
where `/tmp` is tiny or the process may not delete files.

The interpreter is looked up in this order:

1. `--python /path/to/python`
2. `$XLSXTOMYSQL_PYTHON`
3. a venv **next to the binary**: `<exe dir>/../share/xlsxtomysql/venv/…` — this is what
   makes an extracted release tarball or an installed copy work out of the box
4. `~/.local/share/xlsxtomysql/venv/bin/python`
5. `~/.local/bin/python3`, `/usr/bin/python3`, `/usr/local/bin/python3`, `/opt/homebrew/bin/python3`
6. `python3` / `python` on `PATH`; on Windows `.exe` is resolved via `PATHEXT` and the `py` launcher is supported

> **Windows note.** The embedded helper source is ~35 KB, above the ~32 K command-line
> limit of `CreateProcess`, so `python -c <source>` is rejected by the OS. On Windows the
> program therefore writes the helper to a checksummed `.py` under `%TEMP%\xlsxtomysql\`
> and runs that (once per version). Unix keeps using `-c` and never touches the disk.
> To exercise that path on Unix: `XLSXTOMYSQL_PYFILE=1`.

## Quick start

```bash
# headers on row 1, data from row 2, until the end of the sheet
xlsxtomysql students.xlsx Sheet1 students 1 2

# every worksheet of the file, each into its own <sheet-name>.sql
xlsxtomysql students.xlsx

# the 2nd worksheet into a named table; or a whole range of sheets
xlsxtomysql students.xlsx 2 orders
xlsxtomysql students.xlsx "[1-3]"
xlsxtomysql students.xlsx "[1,3,5]"

# batch mode: convert every .xls / .xlsx file in the current directory (all sheets)
xlsxtomysql

# only 300 rows, custom output path, and save the report too
xlsxtomysql big.xlsx Detail orders 2 3 300 --out ./orders.sql --report ./orders.txt

# just look at the inferred column types, generate no SQL
xlsxtomysql students.xlsx Sheet1 students 1 2 --scan-only

# non-standard layout that you have decided to force through, exporting failed rows
xlsxtomysql dirty.xlsx Data tmp 1 2 --force --err-file ./errrows.xlsx
```

Positional arguments, in this fixed order — **every one of them is optional**:

```
xlsxtomysql [FILE.xlsx]  [SHEET]  [TABLE]  [HEADER_ROW]  [FIRST_DATA_ROW]  [ROW_COUNT]
```

| Argument | Meaning | Default when omitted |
|---|---|---|
| `FILE.xlsx` | `.xlsx` (Excel 2007+) or `.xls` (Excel 2003) | omitted → batch mode: every `.xls`/`.xlsx` in the current directory |
| `SHEET` | a worksheet name, a 1-based index (`2` = second sheet), or a range/list expression (`[1-3]`, `[1,3,5]`; brackets optional) | all worksheets |
| `TABLE` | table name, also the default output file name `<TABLE>.sql`; must not be all digits | the (sanitized) sheet name |
| `HEADER_ROW` | 1-based row number holding the column names | `1` |
| `FIRST_DATA_ROW` | 1-based row number of the first data row | `HEADER_ROW + 1` |
| `ROW_COUNT` | number of data rows (`0` also means "to the end") | to the end of the sheet |

Multi-sheet / batch behaviour: each worksheet writes its own `<table>.sql` next to the
source file (`--out` and a custom `TABLE` are not available then); sheets with no data
are skipped with a notice; failed rows from **all** sheets are collected into one
`errrows.xlsx` with one worksheet per source sheet; a "result table 3 - batch summary"
lists every sheet with its success/failed counts and output file.

### A worked example

`examples/students.xlsx` in this repository is a small sheet with a text date column, a
phone column and a percentage column. Running

```bash
xlsxtomysql examples/students.xlsx 学生信息 students 1 2 \
  --out examples/students.sql --report examples/students_report.txt
```

writes `examples/students.sql` (also committed, so you can check the output before installing):

```sql
CREATE TABLE IF NOT EXISTS `students` (
  `姓名` varchar(16),
  `班级` varchar(16),
  `出生日期` date,          -- text in the sheet, recognised as a date
  `身高cm` decimal(4,1),
  `手机号` varchar(32),     -- identifier-like: kept as text so leading zeros survive
  `午餐费` decimal(4,1),
  `出勤率` decimal(4,3),    -- percent-formatted cell, stored as the real ratio
  `备注` varchar(16)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

INSERT INTO `students` (`姓名`,`班级`,`出生日期`,...) VALUES
  ('张小明','三年二班','2017-03-05',128.5,'13800138000',12.5,0.975,'过敏：花生'),
  ...;
```

`examples/students_report.txt` is the console report produced alongside it.

## Command-line options

| Option | Description |
|---|---|
| `--out PATH` | output SQL path (default: `<TABLE>.sql` next to the source file) |
| `--err-file PATH` | failed-rows file (default: `errrows.xlsx` next to the source file) |
| `--report PATH` | also save the console report as plain text |
| `--force` | convert even when a non-standard layout is detected |
| `--scan-only` | scan and print the inferred types only, generate no SQL |
| `--no-hints` / `--hints` | disable / re-enable column-name semantic inference |
| `--full-scan` | force a full scan (slow on large files) |
| `--sample-ratio F` | sampling ratio, default `0.10` |
| `--sample-threshold N` | row count above which sampling kicks in, default `2000` |
| `--sample-min N` / `--sample-max N` | lower/upper bound on sampled rows, default `1000` / `20000` |
| `--seed N` | seed for the sampling start point, for reproducibility |
| `--batch-size N` | rows per `INSERT`, default `200` |
| `--empty-as-null M` | `auto` / `always` / `never`, default `auto` |
| `--varchar-max N` | `varchar` length ceiling, default `1000` |
| `--text-max N` | `text` ceiling, above which `longtext` is used, default `16000` |
| `--max-ident N` | maximum identifier length, default `64` |
| `--merge-scan M` | `auto` / `on` / `off`, default `auto` |
| `--merge-scan-limit N` | maximum file size in MB for merged-cell pre-scan, default `64` |
| `--no-fill-down` | do not fill down merged cells in the data region |
| `--add-id NAME` | append an auto-increment primary key column |
| `--primary-key COL` | use `COL` as the primary key |
| `--drop-table` | emit `DROP TABLE IF EXISTS` |
| `--insert-ignore` | emit `INSERT IGNORE` |
| `--table-comment S` | table comment |
| `--not-null` | add `NOT NULL` to every column |
| `--progress off` | disable the progress bar |
| `--no-color` | disable coloured output |
| `--python PATH` | Python interpreter used to read Excel |
| `--print-errors N` | how many failed rows to print, default `20` |
| `--dump-python` | print the embedded Python helper source and exit |
| `--lang zh\|en` | force the output language (default: auto-detected) |
| `--man` | full manual (option details, type rules, exit codes, limitations), in both languages |
| `-h` / `--help`, `-V` / `--version` | help and version (`--version` includes the target platform) |

## Messy spreadsheets

This is where most Excel-to-SQL tools quietly produce garbage. `xlsxtomysql` reports instead.

| Situation | Behaviour |
|---|---|
| Merged header cells (`A1:D1`) | the value is filled to the right; the report says how many cells were filled |
| Header row inside a vertical merge (`D1:D2`) | the top-left row of the merge is read back, and the merge location is reported |
| Vertically merged data cells (one student spanning 3 rows) | the top-left value fills the whole block; disable with `--no-fill-down` |
| Text dates — `20240101`, `2024-01-01T08:30:00`, `2024/1/1 8:30 AM` | recognised when the column name means a date/time |
| Durations above 24 h — `25:30:00`, `100:00:00` | stored as `time` (MySQL limit 838:59:59) |
| Timestamps with an offset — `2024-01-01T08:30:00+08:00` | stored as `datetime`, with an explicit warning that the offset was dropped |
| Hidden rows / columns | exported normally, with the column names reported |
| Percent formats (displays 12.5 %, stores 0.125) | stores `0.125` and reports how many cells were affected |
| WPS `DISPIMG` / in-cell images | images cannot be written to SQL; the cell becomes empty and this is reported |
| Floating images, comments, hyperlinks, charts | not cell values; ignored without affecting the other columns |
| Ambiguous values — `12/31/2024`, `€1,234.56`, the text `12.5%` | stored as text; the tool does not guess |

**Blank rows** are classified rather than treated as one thing: a run of blank rows is
held back, and only becomes a "blank row inside the data region" — which stops the run
with exit code `2` — if data appears again after it. Blank rows running to the end of the
sheet are trailing and ignored, because Excel exports very often carry hundreds of them.

## Type inference

Types are decided from accumulated per-column statistics, not from the first row:

- text length distribution → `varchar(n)`, `text`, `mediumtext`, `longtext`
- integer magnitude → `tinyint` / `smallint` / `int` / `bigint`
- decimal places → `decimal(p,s)`
- date/time values → `date` / `time` / `datetime`

Column-name semantics refine the result: identifier-like columns (`phone`, `id_card`,
`employee_id`, `bank_card`) are always stored as text so leading zeros survive, while
money-like columns (`amount`, `fee`, `price`) become `decimal`.

Ambiguous input is deliberately **not** guessed. `12/31/2024` could be December 31 or an
invalid March 31 depending on locale, so it stays text.

## Large files and sampling

Above `--sample-threshold` rows (default 2000) the scan switches to a fixed-step jump
sample, and the report states the scan mode and its precision. Two details matter:

- The step is **the prime closest to `1/ratio`** (10 % → every 11th row), not exactly `1/ratio`.
  With a step of 10 and data whose own period is 50 rows, the sampled row numbers only ever
  hit the same 5 residues mod 50 — a whole set of values is systematically missed. (A money
  column whose two-decimal values all live in the other residues gets inferred as one decimal,
  and the conversion then fails on many rows.)
- The first `--sample-min` rows (default 1000) always take part in inference, the number of
  rows read is capped, and `--seed N` makes the sampling start point reproducible.

The progress bar goes to **stderr** (a live bar on a TTY with elapsed/remaining time; a
percentage log when redirected), so it never pollutes a redirected report.

## Language

**Follows the system language by default: Chinese locales get Chinese, everything else
gets English.** Resolution order (first match wins):

| # | Source | Notes |
|---|---|---|
| 1 | `--lang zh` / `--lang en` | explicit override |
| 2 | `XLSXTOMYSQL_LANG` | handy for scripts and CI |
| 3 | `LC_ALL` → `LC_MESSAGES` → `LANG` | a value starting with `zh` means Chinese |
| 4 | platform default | Unix falls back to English (as under the C locale); Windows reads the UI language |

What changes with the language is not only the console: `--help` and `--man` have both
versions, and so do the `.sql` header/footer comments, the `--report` text file, the
`errrows.xlsx` headers (`原行号` / `失败原因` ↔ `Source row` / `Failure reason`), the
ellipsis (… / ...) and list separators in tables.

Rather than hard-coding two sets of output, every user-facing string lives in a `CATALOG`
constant keyed by its Chinese original, and all output boundaries (`Style::wrap`,
`Out::line`, report cells, error construction, helper events) pass through one translation
layer. That makes missed translations detectable by tooling instead of by users:

```bash
python3 tools/i18n_inventory.py          # list every user-facing Chinese string in the source
python3 tools/apply_i18n.py --check      # coverage and placeholder count
python3 tests/check_en_clean.py ./xlsxtomysql   # run every sample in English mode, scan for stray CJK
```

## Cross-platform

| | Linux | macOS | Windows |
|---|---|---|---|
| Binary | ✅ | ✅ | ✅ |
| Console UTF-8 / colour | ✅ native | ✅ native | switches to the UTF-8 code page and enables ANSI escapes at startup |
| Interpreter lookup | `python3` / `python` / venv | also Homebrew paths | `python.exe` via `PATHEXT`, the `py` launcher, `%LOCALAPPDATA%` |
| Home directory | `$HOME` | `$HOME` | `$USERPROFILE` / `%LOCALAPPDATA%` |
| Path joining | `/` | `/` | `\` (always via `PathBuf`) |
| Helper startup | `python -c <source>` (nothing on disk) | same as Linux | writes a temp `.py` and runs it (command-line limit, see above) |
| Local time zone | `localtime_r` | same as Linux | `GetLocalTime` − `GetSystemTime` |
| Broken pipe | restores `SIGPIPE`, `\| head` exits quietly | same as Linux | n/a |

Cross-compiling (the artifact name picks up the target triple automatically):

```bash
rustup target add x86_64-pc-windows-gnu aarch64-unknown-linux-gnu
bash build.sh --release --target x86_64-pc-windows-gnu
bash release.sh --target x86_64-unknown-linux-gnu,aarch64-unknown-linux-gnu,x86_64-pc-windows-gnu
```

**Verified on real hardware:** Linux x86_64 — zero-warning build, the whole test suite and
the byte-for-byte comparison pass. **Reviewed but not executed on real hardware:** the
Windows- and macOS-specific branches. The Windows "write the helper to a temp file" path is
exercised on Linux via `XLSXTOMYSQL_PYFILE=1` and is covered by the automated tests.

## Exit codes

| Code | Meaning |
|---|---|
| `0` | success |
| `1` | runtime error (corrupt or non-Excel file, unreadable sheet, …) |
| `2` | non-standard layout detected, aborted — no SQL written |
| `3` | usage error (missing file, header row after data row, unknown sheet or key column, …) |

## How it is built

```
xlsxtomysql.rs ─┬─ Rust core
                │    argument parsing · layout validation · type inference
                │    SQL generation · report rendering · progress bar · failed-row summary
                │
                └─ const PY_HELPER: &str = r##"…Python helper…"##
                     invoked at runtime as `python3 -c <that source> <mode> <args…>`
                     ├── probe   → file format / worksheet list / dimensions / merges / formulas
                     ├── dump    → stream headers and data rows
                     ├── list    → worksheet names only
                     └── errrows → write failed rows to errrows.xlsx (raw zipfile, no openpyxl)
```

The two processes talk over a **tab-separated line protocol**: the helper's stdout carries
data (`#KIND` / `#SHEET` / `#MERGES` / `#HEADER` / `#ROW` / `#END` / `#OK`, or
`#FATAL<TAB>code<TAB>message`), its stderr carries progress (`#P\t<current>\t<total>`) and
logs. Rust drains the two pipes on **two threads**, so progress refresh never corrupts data
parsing. Cells are encoded as `<kind>:<escaped value>` with a single-character kind
(`e` empty / `i` integer / `d` decimal / `f` scientific / `b` boolean / `s` text /
`D` date / `T` time / `M` datetime / `x` Excel error), classified on the Python side while
reading, with Rust doing the statistics and inference.

The point of the split: the build artifact is a single self-contained binary with no
crates, while Excel parsing — the part where formats are wildly varied and the mature
libraries all live in Python — reuses `openpyxl` / `xlrd` instead of re-implementing OOXML
shared string tables, number formats and the 1900/1904 date bases from scratch.

### Development

```bash
bash build.sh                     # build, zero warnings
bash tests/run_tests.sh           # functional self-test, 112 checks (exit codes, SQL, report text, language, help)
python3 tools/apply_i18n.py --check    # every UI string has an English translation
python3 tools/check_help_sync.py       # every option the code supports is documented in --help / --man
python3 tools/gen_man.py xlsxtomysql.rs [--lang zh]   # regenerate docs/xlsxtomysql*.1 from the source constants
python3 tests/check_en_clean.py ./xlsxtomysql   # run every sample in English mode, scan for stray CJK
```

Tests are self-contained: `tests/make_samples.py` generates the sample workbooks into
`tests/samples/` on first run (needs `openpyxl`; generating the `.xls` samples also needs
`xlwt`), and the English-mode fixtures are generated the same way.

`bash tests/compare_with_python.sh` and `tests/compare_probe.py` are **development-only**
checks that diff this implementation byte-for-byte against an independent Python
implementation of the same tool, which is not part of this repository. If it is not
present they print a notice and exit 0 rather than failing.

`build.sh` verifies on every build that the Python helper embedded in `xlsxtomysql.rs`
matches `py_helper.py`, and prints the diff plus a `--sync` hint if they drifted.

## FAQ

**Do I really need Python installed?**
For reading Excel, yes. The Rust binary itself has no dependencies, but there is no
built-in OOXML/BIFF parser. Swapping the reader for the pure-Rust `calamine` crate is a
one-line change in `build.sh`; the trade-off is giving up "zero crates, one `rustc` command".

**`.xls` says it cannot be read.**
That needs an interpreter with `xlrd`. Install it into any of the venvs listed above.

**Are identifiers quoted?**
Yes. Every identifier is backticked and internal backticks are doubled, so keywords and
spaces in column names are safe.

**Why is the progress bar sometimes not animated?**
When output is redirected to a file or a pipe it degrades to percentage log lines to avoid
flooding the log. In a terminal you get the animated bar.

**Does it drop any rows?**
No. Rows that cannot be converted are counted, listed with a reason, and exported to
`errrows.xlsx`.

## License

MIT — see [LICENSE](LICENSE). The embedded Python helper is under the same licence.

## Keywords

Excel to MySQL · xlsx to SQL · xls to MySQL · Excel to SQL converter · spreadsheet to MySQL ·
generate CREATE TABLE from Excel · generate INSERT statements from Excel · xlsx to MySQL DDL ·
Excel 导入 MySQL · Excel 转 SQL · 表格转数据库 · 建表语句生成 · xlsx2mysql ·
command-line tool · CLI · single binary · cross-platform · Linux · macOS · Windows · Arch Linux · AUR
