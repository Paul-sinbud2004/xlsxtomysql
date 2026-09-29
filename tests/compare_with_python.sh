#!/usr/bin/env bash
# 双版本对照测试：同一份输入，分别用 Python 实现和本仓库生成 SQL，
# 逐字节比对 CREATE TABLE 与 INSERT 部分。
#
#   bash tests/compare_with_python.sh              # 用默认 Python 解释器
#   PY=/path/to/python bash tests/compare_with_python.sh
#
# 注意：这是**开发期**检查，需要一个并行的 Python 实现（默认找 ../xlsxtomysql/xlsxtomysql.py）。
# 本仓库不含该实现；找不到时会明确提示并跳过，不影响本仓库自身的自测。
set -u

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(dirname "$HERE")"                 # 仓库根
WS="$(dirname "$ROOT")"                   # 工作目录（可能没有 Python 实现）
PY_VERSION="${PY_VERSION:-$WS/xlsxtomysql/xlsxtomysql.py}"
SAMPLES="${SAMPLES:-$HERE/samples}"
RS="$ROOT/xlsxtomysql"
OUT="$HERE/cmp"

PY="${PY:-$HOME/.local/share/xlsxtomysql/venv/bin/python}"
[ -x "$PY" ] || PY=python3

if [ ! -f "$PY_VERSION" ]; then
  echo "== 跳过：未找到 Python 实现（$PY_VERSION）=="
  echo "   这个是开发期对照检查，公开发布版仓库不含 Python 实现。"
  echo "   想跑的话：PY_VERSION=/path/to/xlsxtomysql.py bash tests/compare_with_python.sh"
  exit 0
fi

mkdir -p "$OUT"
[ -x "$RS" ] || { echo "先编译: bash build.sh"; exit 1; }

# 样本不存在时自动生成
if [ ! -d "$SAMPLES" ] || [ -z "$(ls -A "$SAMPLES" 2>/dev/null)" ]; then
  echo "== 生成测试样本到 $SAMPLES =="
  "$PY" "$HERE/make_samples.py" || { echo "× 样本生成失败"; exit 1; }
fi

pass=0; fail=0

# 只取建表语句与 INSERT 语句，忽略两版各自的头部注释
extract() {
  awk '/^CREATE TABLE/,/^\) ENGINE/ {print} /^INSERT/ {print}' "$1"
}

cmp_case() {
  local desc="$1"; shift
  "$PY" "$PY_VERSION" "$@" --progress off --no-color --out "$OUT/py.sql" >/dev/null 2>&1
  local pcode=$?
  "$RS" "$@" --progress off --no-color --out "$OUT/rs.sql" >/dev/null 2>&1
  local rcode=$?
  extract "$OUT/py.sql" > "$OUT/py.x" 2>/dev/null
  extract "$OUT/rs.sql" > "$OUT/rs.x" 2>/dev/null

  if [ "$pcode" != "$rcode" ]; then
    printf "  [FAIL] %-32s 退出码不同 py=%s rs=%s\n" "$desc" "$pcode" "$rcode"
    fail=$((fail+1)); return
  fi
  if diff -q "$OUT/py.x" "$OUT/rs.x" >/dev/null 2>&1; then
    printf "  [PASS] %-32s exit=%s\n" "$desc" "$rcode"
    pass=$((pass+1))
  else
    printf "  [FAIL] %-32s 输出不同 (exit=%s)\n" "$desc" "$rcode"
    diff "$OUT/py.x" "$OUT/rs.x" | head -14 | sed 's/^/         /'
    fail=$((fail+1))
  fi
}

echo "== 1. xlsx 基础 =="
cmp_case "标准表 500 行"        "$SAMPLES/01_normal.xlsx" 学生信息 t1 1 2
cmp_case "合并表头"             "$SAMPLES/03_merged_header.xlsx" 数据 t3 2 3
cmp_case "单列合并表头"         "$SAMPLES/09_merged_single.xlsx" 数据 t9 1 2
cmp_case "多工作表·第二张"      "$SAMPLES/08_multi_sheet.xlsx" 第二张表 t8 1 2

echo "== 2. 非标准格式（--force 后转换 + 失败行）=="
cmp_case "分隔行/空行/重复表头"  "$SAMPLES/02_messy.xlsx" 数据 t2 1 2 --force
cmp_case "脏数据表"             "$SAMPLES/04_dirty.xlsx" 数据 t4 1 2 --force

echo "== 3. 大数据 + 抽样跳跃扫描 =="
cmp_case "30000 行 · 抽样"       "$SAMPLES/05_big.xlsx" bigdata t5 1 2 --seed 7
cmp_case "30000 行 · 全量"       "$SAMPLES/05_big.xlsx" bigdata t5b 1 2 --full-scan

echo "== 4. Excel 2003 (.xls) =="
cmp_case ".xls 普通表"          "$SAMPLES/06_legacy.xls" 旧数据 t6 1 2
cmp_case ".xls 合并单元格"       "$SAMPLES/11_xls_merge.xls" 旧表 t10 1 2

echo "== 5. 字段名清洗 =="
cmp_case "空白/重名/超长/非法字符" "$SAMPLES/07_names.xlsx" 怪字段名 t7 1 2

echo "== 6. 空值策略 =="
cmp_case "auto"                 "$SAMPLES/12_empty.xlsx" T t11 1 2
cmp_case "always"               "$SAMPLES/12_empty.xlsx" T t11a 1 2 --empty-as-null always
cmp_case "never"                "$SAMPLES/12_empty.xlsx" T t11n 1 2 --empty-as-null never

echo "== 7. 输出选项 =="
cmp_case "自增主键+表注释+NOT NULL" "$SAMPLES/01_normal.xlsx" 学生信息 t13 1 2 \
         --add-id id --table-comment "测试表" --not-null --drop-table
cmp_case "指定主键 + INSERT IGNORE" "$SAMPLES/01_normal.xlsx" 学生信息 t15 1 2 \
         --primary-key 姓名 --insert-ignore
cmp_case "批量 50"               "$SAMPLES/01_normal.xlsx" 学生信息 t16 1 2 --batch-size 50
cmp_case "限制 300 行"           "$SAMPLES/01_normal.xlsx" 学生信息 t14 1 2 300
cmp_case "关闭字段名语义推断"      "$SAMPLES/01_normal.xlsx" 学生信息 t17 1 2 --no-hints

echo
echo "================ 通过 $pass / 失败 $fail ================"
[ "$fail" = "0" ]
