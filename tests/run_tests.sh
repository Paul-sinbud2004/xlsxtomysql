#!/usr/bin/env bash
# xlsxtomysql 功能自测
#   bash tests/run_tests.sh
#
# 自包含：测试样本由 tests/make_samples.py 生成到 tests/samples/，首次运行会自动生成。
# 需要 openpyxl；生成 .xls 样本还需要 xlwt。用 PY=<python> 指定解释器。
set -u

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(dirname "$HERE")"

BIN="${BIN:-$ROOT/xlsxtomysql}"
SAMPLES="${SAMPLES:-$HERE/samples}"
GEN="$HERE/make_samples.py"
PY="${PY:-$HOME/.local/share/xlsxtomysql/venv/bin/python}"
[ -x "$PY" ] || PY=python3
D="$SAMPLES"
TMP="$ROOT/tests/tmp"
mkdir -p "$TMP"

[ -x "$BIN" ] || { echo "先编译: bash build.sh"; exit 1; }

# 样本不存在或为空时自动生成
if [ ! -d "$D" ] || [ -z "$(ls -A "$D" 2>/dev/null)" ]; then
  echo "== 生成测试样本到 $D =="
  "$PY" "$GEN" || {
    echo "× 样本生成失败：需要一个装有 openpyxl 的 Python（生成 .xls 样本还需 xlwt）。"
    echo "  用 PY=/path/to/python bash tests/run_tests.sh 指定解释器。"
    exit 1
  }
fi
[ -d "$D" ] || { echo "缺少测试样本目录 $D"; exit 1; }

pass=0; fail=0; n=0

run() {   # run <期望退出码> <说明> <日志文件> <参数...>
  local want="$1" desc="$2" logf="$3"; shift 3
  n=$((n+1))
  local out code
  out=$("$BIN" "$@" --progress off --no-color 2>&1)
  code=$?
  printf '%s\n' "$out" > "$logf"
  if [ "$code" = "$want" ]; then
    echo "  [PASS] $desc  (exit=$code)"
    pass=$((pass+1))
  else
    echo "  [FAIL] $desc  期望 exit=$want 实际 exit=$code"
    echo "$out" | tail -12 | sed 's/^/         /'
    fail=$((fail+1))
  fi
}

expect_grep() {  # expect_grep <说明> <输出文件> <正则>
  if grep -qE "$3" "$2" 2>/dev/null; then
    echo "  [PASS] $1"
    pass=$((pass+1))
  else
    echo "  [FAIL] $1 —— 输出中未找到 /$3/"
    fail=$((fail+1))
  fi
}

echo "== 1. 标准表 (.xlsx) =="
run 0 "标准表转换" "$TMP/1.txt" "$D/01_normal.xlsx" 学生信息 t_normal 1 2
expect_grep "生成建表语句 utf8mb4" "$D/t_normal.sql" "DEFAULT CHARSET=utf8mb4"
expect_grep "出生日期 → date"      "$D/t_normal.sql" '`出生日期` date'
expect_grep "是否住校 → tinyint(1)" "$D/t_normal.sql" '`是否住校` tinyint\(1\)'
expect_grep "日期字面量加引号"      "$D/t_normal.sql" "'2018-01-01'"
expect_grep "成功 500 行"            "$TMP/1.txt" "成功行数.*500"

echo "== 2. 非标准格式（应阻断） =="
run 2 "分隔行/重复表头 → 阻断" "$TMP/2.txt" "$D/02_messy.xlsx" 数据 t_messy 1 2
expect_grep "指出空行位置"     "$TMP/2.txt" "第 12 行.*分隔用的空行"
expect_grep "指出分隔线位置"   "$TMP/2.txt" "第 23 行.*分隔线"
expect_grep "指出重复表头位置" "$TMP/2.txt" "第 29 行.*重复的字段名行"

echo "== 3. --force 强行转换 + 失败行导出 =="
rm -f "$D/errrows.xlsx"
run 0 "--force 继续转换" "$TMP/3.txt" "$D/02_messy.xlsx" 数据 t_messy 1 2 --force
expect_grep "失败行含原行号" "$TMP/3.txt" "23.*不是整数"
test -f "$D/errrows.xlsx" && { echo "  [PASS] errrows.xlsx 已生成"; pass=$((pass+1)); } \
                         || { echo "  [FAIL] errrows.xlsx 未生成"; fail=$((fail+1)); }
# 校验 errrows.xlsx 的表头与内容
if "$PY" "$ROOT/tests/check_errrows.py" "$D/errrows.xlsx"; then
  echo "  [PASS] errrows.xlsx 内容可读、含原行号与失败原因"
  pass=$((pass+1))
else
  echo "  [FAIL] errrows.xlsx 内容校验失败"
  fail=$((fail+1))
fi

echo "== 4. 合并单元格表头 =="
run 0 "两行表头（字段名在第 2 行）" "$TMP/4.txt" "$D/03_merged_header.xlsx" 数据 t_merged 2 3
expect_grep "建表字段正确" "$D/t_merged.sql" '`午餐费`'
run 0 "字段名行内合并单元格自动填充" "$TMP/4b.txt" "$D/09_merged_single.xlsx" 数据 t_msingle 1 2
expect_grep "横向合并已填充"   "$TMP/4b.txt" "处合并单元格，已自动填充"
expect_grep "合并列重名去重"   "$D/t_msingle.sql" '`基本信息_2`'

echo "== 5. 怪字段名清洗 =="
run 0 "字段名清洗" "$TMP/5.txt" "$D/07_names.xlsx" 怪字段名 t_names 1 2
expect_grep "重名自动加后缀" "$TMP/5.txt" "姓名_2"
expect_grep "空字段名自动命名" "$TMP/5.txt" "col_d"
expect_grep "超长截断"       "$TMP/5.txt" "超长截断"
expect_grep "非法字符替换"   "$TMP/5.txt" "a_b_c"

echo "== 6. 大数据量 + 抽样跳跃扫描 =="
run 0 "30000 行抽样扫描" "$TMP/6.txt" "$D/05_big.xlsx" bigdata t_big 1 2 --seed 7
expect_grep "启用抽样跳跃扫描" "$TMP/6.txt" "跳跃抽样"
expect_grep "成功 30000 行"    "$TMP/6.txt" "成功行数.*30,000"

echo "== 7. Excel 2003 (.xls) =="
run 0 ".xls 格式转换" "$TMP/7.txt" "$D/06_legacy.xls" 旧数据 t_legacy 1 2
expect_grep "识别为 xls"   "$TMP/7.txt" "Excel 2003"
expect_grep "日期列 → date" "$D/t_legacy.sql" '`出生日期` date'
run 0 ".xls 合并单元格" "$TMP/7b.txt" "$D/11_xls_merge.xls" 旧表 t_xlsm 1 2
expect_grep "xls 合并已填充" "$TMP/7b.txt" "处合并单元格，已自动填充"
expect_grep "xls 合并列重名去重" "$D/t_xlsm.sql" '`基本信息_2`'

echo "== 8. 空值策略 =="
run 0 "auto" "$TMP/12a.txt" "$D/12_empty.xlsx" T t_e1 1 2 --empty-as-null auto
expect_grep "auto: 文本空 → ''" "$D/t_e1.sql" ",'',"
expect_grep "auto: 数值空 → NULL" "$D/t_e1.sql" "乙',NULL"
run 0 "always" "$TMP/12b.txt" "$D/12_empty.xlsx" T t_e2 1 2 --empty-as-null always
expect_grep "always: 空值一律 NULL" "$D/t_e2.sql" "\('甲',90,NULL"
run 0 "never" "$TMP/12c.txt" "$D/12_empty.xlsx" T t_e3 1 2 --empty-as-null never
expect_grep "never: 空值一律 ''" "$D/t_e3.sql" ",'',"

echo "== 9. 异常输入 =="
printf 'a,b\n1,2\n' > "$TMP/fake.xlsx"
run 1 "伪装成 xlsx 的文本文件" "$TMP/13a.txt" "$TMP/fake.xlsx" S t 1 2
expect_grep "给出格式提示" "$TMP/13a.txt" "无法识别的文件格式"
run 3 "共几行为负数" "$TMP/13b.txt" "$D/01_normal.xlsx" 学生信息 t_x 1 2 -5
run 0 "共几行=0 表示到文件尾" "$TMP/13c.txt" "$D/01_normal.xlsx" 学生信息 t_x0 1 2 0
expect_grep "仍是 500 行" "$TMP/13c.txt" "共多少行数据.*500"
run 0 "共几行超出范围自动截断" "$TMP/13d.txt" "$D/01_normal.xlsx" 学生信息 t_x 1 2 999999
expect_grep "提示已截断" "$TMP/13d.txt" "超出工作表范围"

echo "== 10. 多工作表 =="
run 0 "指定第二张工作表" "$TMP/8.txt" "$D/08_multi_sheet.xlsx" 第二张表 t_sheet2 1 2
expect_grep "单价 → decimal" "$D/t_sheet2.sql" '`单价` decimal'
run 3 "工作表不存在 → 提示可选表" "$TMP/8b.txt" "$D/08_multi_sheet.xlsx" 不存在的表 t_x 1 2
expect_grep "列出可用工作表" "$TMP/8b.txt" "可用工作表.*第一张表"

echo "== 11. 参数与异常 =="
run 3 "行号参数颠倒" "$TMP/9a.txt" "$D/01_normal.xlsx" 学生信息 t_x 5 3
run 3 "文件不存在"   "$TMP/9b.txt" "/no/such/file.xlsx" S t_x 1 2
run 3 "表头行无内容" "$TMP/9c.txt" "$D/01_normal.xlsx" 学生信息 t_x 600 601
run 0 "限制行数 共300行" "$TMP/9d.txt" "$D/01_normal.xlsx" 学生信息 t_limit 1 2 300
expect_grep "共300行" "$TMP/9d.txt" "共多少行数据.*300"
run 0 "--scan-only 不生成 sql" "$TMP/9e.txt" "$D/01_normal.xlsx" 学生信息 t_scan 1 2 --scan-only
expect_grep "仅扫描提示" "$TMP/9e.txt" "仅扫描模式"

echo "== 12. 输出选项 =="
run 0 "--drop-table / --add-id / 主键 / 注释" "$TMP/10.txt" "$D/01_normal.xlsx" 学生信息 t_opt 1 2 \
     --drop-table --add-id id --table-comment "测试表" --not-null
expect_grep "DROP TABLE" "$D/t_opt.sql" "DROP TABLE IF EXISTS"
expect_grep "自增主键"   "$D/t_opt.sql" "AUTO_INCREMENT"
expect_grep "表注释"     "$D/t_opt.sql" "COMMENT='测试表'"
expect_grep "NOT NULL"   "$D/t_opt.sql" "NOT NULL"
# 自增列后不应出现双逗号
if grep -qE "AUTO_INCREMENT,," "$D/t_opt.sql"; then
  echo "  [FAIL] 自增主键列后出现多余逗号"; fail=$((fail+1))
else
  echo "  [PASS] 自增主键列语法正确（无多余逗号）"; pass=$((pass+1))
fi
run 0 "--primary-key + --insert-ignore" "$TMP/10b.txt" "$D/01_normal.xlsx" 学生信息 t_pk 1 2 \
     --primary-key 姓名 --insert-ignore
expect_grep "INSERT IGNORE" "$D/t_pk.sql" "INSERT IGNORE"
expect_grep "指定主键"       "$D/t_pk.sql" 'PRIMARY KEY \(`姓名`\)'
run 3 "主键名不存在 → 报错" "$TMP/10c.txt" "$D/01_normal.xlsx" 学生信息 t_pk2 1 2 --primary-key 不存在的列

echo "== 13. 导出报表 / 批次大小 =="
run 0 "--report + --batch-size 50" "$TMP/11.txt" "$D/01_normal.xlsx" 学生信息 t_rep 1 2 \
     --report "$TMP/report.txt" --batch-size 50
test -f "$TMP/report.txt" && { echo "  [PASS] 报表文件已生成"; pass=$((pass+1)); } \
                          || { echo "  [FAIL] 报表文件未生成"; fail=$((fail+1)); }
cnt=$(grep -c "^INSERT INTO" "$D/t_rep.sql")
if [ "$cnt" = "10" ]; then echo "  [PASS] 500 行 / 批量 50 = 10 条 INSERT"; pass=$((pass+1));
else echo "  [FAIL] INSERT 条数=$cnt，期望 10"; fail=$((fail+1)); fi

echo "== 14. 特殊表格边界（回归） =="
# 文本日期：ISO T 分隔 / 紧凑 / 斜杠 / 毫秒 / 上下午 / 超 24 小时时长
run 0 "文本日期多样写法" "$TMP/14a.txt" "$D/13_date_text.xlsx" 数据 t_dt 1 2
expect_grep "ISO T 日期识别为 datetime" "$TMP/14a.txt" "登记时间.*datetime"
expect_grep "紧凑日期识别为 date"       "$TMP/14a.txt" "出生日期.*date"
expect_grep "超 24 小时识别为 time"     "$TMP/14a.txt" "加班时长.*time"
# 数据区纵向合并：一人占 3 行，空单元格应按左上角补齐
run 0 "数据区纵向合并填充" "$TMP/14b.txt" "$D/14_merge_vertical.xlsx" 数据 t_mg 1 2
expect_grep "合并单元格已补齐" "$TMP/14b.txt" "数据区合并单元格已补齐"
cnt=$(grep -c "'张三'" "$D/t_mg.sql")
if [ "$cnt" = "3" ]; then echo "  [PASS] 张三补齐到 3 行"; pass=$((pass+1));
else echo "  [FAIL] 张三出现 $cnt 次，期望 3"; fail=$((fail+1)); fi
# 字段名行落在纵向合并区内：字段名在合并区左上角那一行
run 0 "字段名行在纵向合并区内" "$TMP/14c.txt" "$D/15_header_vmerge.xlsx" 数据 t_hv 2 3
expect_grep "回读合并区左上角作为字段名" "$TMP/14c.txt" "位于纵向合并区"
expect_grep "字段名正确" "$D/t_hv.sql" "语文. tinyint"
# 数据区中间空行：属于非标准格式，应阻断
run 2 "数据区中间空行 → 阻断" "$TMP/14d.txt" "$D/16_blank_mid.xlsx" 数据 t_bm 1 2
expect_grep "指出空行位置" "$TMP/14d.txt" "第 3 行.*分隔用的空行"
# 隐藏列 + 百分比格式：都要给出提示
run 0 "隐藏列 + 百分比格式" "$TMP/14e.txt" "$D/17_hidden_pct.xlsx" 数据 t_hp 1 2
expect_grep "隐藏列提示" "$TMP/14e.txt" "检测到 1 个隐藏列"
expect_grep "百分比格式提示" "$TMP/14e.txt" "百分比格式的单元格"
# 带时区偏移的日期时间：偏移会被丢弃，须提示
run 0 "带时区偏移的日期时间" "$TMP/14f.txt" "$D/18_tz_offset.xlsx" 数据 t_tz 1 2
expect_grep "时区偏移被丢弃提示" "$TMP/14f.txt" "带时区偏移"
expect_grep "按 datetime 入库" "$D/t_tz.sql" "上传时间. datetime"

echo "== 15. 内嵌 Python 助手 =="
"$BIN" --dump-python > "$TMP/dump.py" 2>/dev/null
if [ -s "$TMP/dump.py" ] && "$PY" -c "import ast,sys; ast.parse(open('$TMP/dump.py',encoding='utf-8').read())"; then
  echo "  [PASS] --dump-python 导出的助手源码语法正确（$(wc -l < "$TMP/dump.py") 行）"
  pass=$((pass+1))
else
  echo "  [FAIL] --dump-python 导出的源码无法解析"; fail=$((fail+1))
fi

echo "== 16. 语言 / 帮助 / 手册 / 跨平台 =="

# 英文断言必须用英文字段名的样例：中文样例里的「姓名/出生日期」是数据，
# 出现在报表和 .sql 里是正常的，不能算残留。
FIX="$HERE/enfix"
if [ ! -f "$FIX/en_clean.xlsx" ]; then
  "$PY" "$HERE/make_en_fixture.py" "$FIX" > /dev/null 2>&1
fi
if [ ! -f "$FIX/en_clean.xlsx" ]; then
  echo "  [SKIP] 缺少英文样例（$PY $HERE/make_en_fixture.py $FIX）"
fi
EC=("$FIX/en_clean.xlsx" Data)

# 判定一份输出里有没有汉字（英文模式不允许出现）
has_cjk() {
  "$PY" - "$1" <<'PYEOF'
import re, sys
t = open(sys.argv[1], encoding="utf-8", errors="replace").read()
sys.exit(0 if re.search(r"[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]", t) else 1)
PYEOF
}
ok() { echo "  [PASS] $1"; pass=$((pass+1)); }
no() { echo "  [FAIL] $1"; fail=$((fail+1)); }

CNV=("$D/01_normal.xlsx" 学生信息)

# --- 16a 默认跟随系统语言：中文环境 → 中文 ---
env LANG=zh_CN.UTF-8 LC_ALL= LC_MESSAGES= XLSXTOMYSQL_LANG= \
  "$BIN" "${CNV[0]}" "${CNV[1]}" t_lang_zh 1 2 --out "$TMP/lang_zh.sql" \
  --progress off --no-color > "$TMP/16a.txt" 2>&1
if grep -q "预扫描" "$TMP/16a.txt"; then ok "LANG=zh_CN 时输出中文"; else no "LANG=zh_CN 时未输出中文"; fi

# --- 16b 默认跟随系统语言：C / 英文环境 → 英文且无中文残留 ---
env LANG=C LC_ALL=C XLSXTOMYSQL_LANG= \
  "$BIN" "${EC[0]}" "${EC[1]}" t_lang_en 1 2 --out "$TMP/lang_en.sql" \
  --progress off --no-color > "$TMP/16b.txt" 2>&1
if grep -qi "scan" "$TMP/16b.txt"; then ok "LANG=C 时输出英文"; else no "LANG=C 时未输出英文"; fi
if has_cjk "$TMP/16b.txt"; then no "LANG=C 时控制台仍有中文"; else ok "LANG=C 时控制台无中文残留"; fi
if has_cjk "$TMP/lang_en.sql"; then no "英文模式下 .sql 注释仍有中文"; else ok "英文模式下 .sql 注释无中文"; fi

# --- 16c XLSXTOMYSQL_LANG 覆盖系统 locale ---
env LANG=C LC_ALL=C XLSXTOMYSQL_LANG=zh \
  "$BIN" "${CNV[0]}" "${CNV[1]}" t_lang_v 1 2 --out "$TMP/lang_v.sql" \
  --progress off --no-color > "$TMP/16c.txt" 2>&1
if grep -q "预扫描" "$TMP/16c.txt"; then ok "XLSXTOMYSQL_LANG=zh 覆盖 LANG=C"; else no "XLSXTOMYSQL_LANG 未生效"; fi

# --- 16d --lang 覆盖环境变量 ---
env XLSXTOMYSQL_LANG=zh \
  "$BIN" "${EC[0]}" "${EC[1]}" t_lang_p 1 2 --lang=en --out "$TMP/lang_p.sql" \
  --progress off --no-color > "$TMP/16d.txt" 2>&1
if has_cjk "$TMP/16d.txt"; then no "--lang=en 未覆盖环境变量"; else ok "--lang=en 覆盖 XLSXTOMYSQL_LANG=zh"; fi

# --- 16e --lang 也影响 --report 与 errrows 表头 ---
env LANG=C LC_ALL=C \
  "$BIN" "${EC[0]}" "${EC[1]}" t_lang_r 1 2 --lang=en \
  --out "$TMP/lang_r.sql" --report "$TMP/lang_r.txt" --progress off --no-color > /dev/null 2>&1
if [ -s "$TMP/lang_r.txt" ] && ! has_cjk "$TMP/lang_r.txt"; then
  ok "--report 报表跟随语言（英文无中文）"
else
  no "--report 报表未跟随语言"
fi

# --- 16f --help / --man：中英各一份，退出码 0，且不依赖位置参数 ---
"$BIN" --lang=zh --help > "$TMP/16f_zh.txt" 2>&1; c=$?
if [ "$c" = 0 ] && grep -q "用法" "$TMP/16f_zh.txt" && grep -q -- "--man" "$TMP/16f_zh.txt"; then
  ok "--help 中文（退出码 0）"
else
  no "--help 中文异常（exit=$c）"
fi
"$BIN" --lang=en --help > "$TMP/16f_en.txt" 2>&1; c=$?
if [ "$c" != 0 ] || ! grep -q "Usage" "$TMP/16f_en.txt"; then
  no "--help 英文异常（exit=$c）"
elif has_cjk "$TMP/16f_en.txt"; then
  no "--help 英文里混进了中文"
else
  ok "--help 英文（退出码 0，无中文）"
fi
"$BIN" --lang=zh --man > "$TMP/16g_zh.txt" 2>&1; c=$?
if [ "$c" = 0 ] && grep -q "退出码" "$TMP/16g_zh.txt" && grep -q "环境变量" "$TMP/16g_zh.txt"; then
  ok "--man 中文手册完整（退出码 0）"
else
  no "--man 中文手册异常（exit=$c）"
fi
"$BIN" --lang=en --man > "$TMP/16g_en.txt" 2>&1; c=$?
if [ "$c" = 0 ] && grep -q "Exit codes" "$TMP/16g_en.txt" && grep -q "Environment variables" "$TMP/16g_en.txt"; then
  ok "--man 英文手册完整（退出码 0）"
else
  no "--man 英文手册异常（exit=$c）"
fi
# 手册里必须覆盖代码支持的选项（防止加了选项忘了写文档）
if "$PY" "$ROOT/tools/check_help_sync.py" "$ROOT/xlsxtomysql.rs" > "$TMP/16h.txt" 2>&1; then
  ok "help/man 与代码选项同步（$(grep -o '选项: [0-9]* 个' "$TMP/16h.txt" | head -1)）"
else
  no "help/man 有选项未写进文档"
  sed 's/^/         /' "$TMP/16h.txt" | tail -6
fi

# --- 16i 管道截断：--man | head 不能 panic（恢复 SIGPIPE） ---
"$BIN" --man 2> "$TMP/16i.err" | head -2 > /dev/null
if grep -qi "panic" "$TMP/16i.err"; then
  no "管道被下游关闭时出现 panic"
else
  ok "管道截断时安静退出（无 panic）"
fi

# --- 16j 参数与路径错误的退出码 ---
"$BIN" $D/01_normal.xlsx 学生信息 t_e 1 2 --nosuch > "$TMP/16j1.txt" 2>&1; c=$?
[ "$c" = 3 ] && ok "未知选项 → exit 3" || no "未知选项退出码 $c（期望 3）"
"$BIN" > "$TMP/16j2.txt" 2>&1; c=$?
[ "$c" = 3 ] && ok "无参数 → exit 3 且提示 --help" || no "无参数退出码 $c（期望 3）"
env LANG=C LC_ALL=C \
  "$BIN" "${CNV[0]}" "${CNV[1]}" t_e2 1 2 --out "$TMP/nodir/x.sql" --no-color > "$TMP/16j3.txt" 2>&1; c=$?
if [ "$c" = 3 ] && grep -qi "directory" "$TMP/16j3.txt"; then
  ok "输出目录不存在 → exit 3 且有明确提示"
else
  no "输出目录不存在时提示不明确（exit=$c）"
fi

# --- 16k 原子输出：失败不留半截 .sql，也不覆盖上一次的成果 ---
"$BIN" "${CNV[0]}" "${CNV[1]}" t_atom 1 2 --out "$TMP/atom.sql" --progress off --no-color > /dev/null 2>&1
BEFORE=$(cksum < "$TMP/atom.sql" 2>/dev/null)
"$BIN" "${CNV[0]}" 不存在的工作表 t_atom 1 2 --out "$TMP/atom.sql" --progress off --no-color > /dev/null 2>&1
AFTER=$(cksum < "$TMP/atom.sql" 2>/dev/null)
if [ -n "$BEFORE" ] && [ "$BEFORE" = "$AFTER" ]; then
  ok "转换失败后原有的 .sql 未被破坏"
else
  no "转换失败破坏了上一次生成的 .sql"
fi
if find "$TMP" -maxdepth 1 -name '*.part' | grep -q .; then
  no "留下了未清理的 .part 临时文件"
else
  ok "没有残留 .part 临时文件"
fi

# --- 16l 助手启动方式：命令行 vs 临时脚本（Windows 走后者）---
env XLSXTOMYSQL_PYFILE=1 LANG=zh_CN.UTF-8 LC_ALL= \
  "$BIN" "${CNV[0]}" "${CNV[1]}" t_lang_zh 1 2 --out "$TMP/pyfile.sql" \
  --progress off --no-color > /dev/null 2>&1; c=$?
if [ "$c" = 0 ] && [ -s "$TMP/pyfile.sql" ]; then
  ok "助手落临时脚本方式可正常转换（模拟 Windows 路径）"
else
  no "临时脚本方式转换失败（exit=$c）"
fi
# 头尾注释里有时间戳，先剔掉再比 SQL 正文
grep -v "^-- " "$TMP/lang_zh.sql" > "$TMP/cmp_cmd.txt"
grep -v "^-- " "$TMP/pyfile.sql" > "$TMP/cmp_file.txt"
if diff -q "$TMP/cmp_cmd.txt" "$TMP/cmp_file.txt" > /dev/null 2>&1; then
  ok "两种助手启动方式产物一致"
else
  no "两种助手启动方式产物不一致"
  diff "$TMP/cmp_cmd.txt" "$TMP/cmp_file.txt" | head -6 | sed 's/^/         /'
fi

# --- 16m 译文覆盖率与 man page 同源 ---
if "$PY" "$ROOT/tools/apply_i18n.py" --check > "$TMP/16m1.txt" 2>&1; then
  ok "译文覆盖率检查通过（$(grep -o '译文表条目: [0-9]*' "$TMP/16m1.txt" | head -1)）"
else
  no "有中文文案没有对应英文"
  sed 's/^/         /' "$TMP/16m1.txt" | tail -8
fi
"$PY" "$ROOT/tools/gen_man.py" "$ROOT/xlsxtomysql.rs" -o "$TMP/gen.1" > /dev/null 2>&1
if diff -q "$TMP/gen.1" "$ROOT/docs/xlsxtomysql.1" > /dev/null 2>&1; then
  ok "man page 与 --man 手册内容同源"
else
  no "man page 与手册已不同步（跑 tools/gen_man.py 重新生成）"
fi


echo "== 17. 多 sheet 缺省参数 / 批量模式 =="

# 不指定 sheet → 转全部，每 sheet 一个 <sheet名>.sql；无数据的「说明」表被跳过
rm -f "$D/第一张表.sql" "$D/第二张表.sql" "$D/说明.sql"
run 0 "不指定 sheet → 转全部" "$TMP/17a.txt" "$D/08_multi_sheet.xlsx"
[ -f "$D/第一张表.sql" ] && ok "生成 第一张表.sql" || no "生成 第一张表.sql"
[ -f "$D/第二张表.sql" ] && ok "生成 第二张表.sql" || no "生成 第二张表.sql"
expect_grep "无数据的说明表被跳过" "$TMP/17a.txt" "已跳过该工作表"

# 编号 / 区间 / 列表
rm -f "$D/第一张表.sql" "$D/第二张表.sql"
run 0 "编号选择第 2 个 sheet" "$TMP/17b.txt" "$D/08_multi_sheet.xlsx" 2
{ [ -f "$D/第二张表.sql" ] && [ ! -f "$D/第一张表.sql" ]; } \
  && ok "编号 2 只生成第二张表" || no "编号 2 只生成第二张表"
run 0 "区间 [1-2]" "$TMP/17c.txt" "$D/08_multi_sheet.xlsx" "[1-2]"
{ [ -f "$D/第一张表.sql" ] && [ -f "$D/第二张表.sql" ]; } \
  && ok "区间 [1-2] 生成两个 sql" || no "区间 [1-2] 生成两个 sql"
rm -f "$D/第二张表.sql"
run 0 "列表 [1,3]" "$TMP/17d.txt" "$D/08_multi_sheet.xlsx" "[1,3]"
{ [ -f "$D/第一张表.sql" ] && [ ! -f "$D/第二张表.sql" ]; } \
  && ok "列表 [1,3] 只生成第一张表" || no "列表 [1,3] 只生成第一张表"

# 缺省表名（文件 + sheet 两个参数）与缺省行号（1 / 2 / 到文件尾）
rm -f "$D/第二张表.sql"
run 0 "指定 sheet、缺省表名" "$TMP/17e.txt" "$D/08_multi_sheet.xlsx" 第二张表
[ -f "$D/第二张表.sql" ] && ok "缺省表名 = sheet 名" || no "缺省表名 = sheet 名"
run 0 "缺省行号等价旧写法" "$TMP/17f.txt" "$D/08_multi_sheet.xlsx" 第二张表 t_s2
expect_grep "缺省行号结果一致" "$D/t_s2.sql" '`单价` decimal'

# 冲突与错误
run 3 "多 sheet + --out → 用法错误" "$TMP/17g.txt" "$D/08_multi_sheet.xlsx" --out "$TMP/x17.sql"
run 3 "多 sheet + 表名 → 用法错误" "$TMP/17h.txt" "$D/08_multi_sheet.xlsx" t_x
run 3 "编号越界 → 用法错误" "$TMP/17i.txt" "$D/08_multi_sheet.xlsx" 99
expect_grep "越界提示包含表数量" "$TMP/17i.txt" "共 3 个工作表"

# errrows 多 sheet：两个 sheet 各有失败行（--seed 5 保证脏行不被抽样判型）
"$PY" - "$TMP/dirty_multi.xlsx" <<'PY17GEN'
import sys
from openpyxl import Workbook
wb = Workbook(); wb.remove(wb.active)
for name in ("脏甲", "脏乙"):
    ws = wb.create_sheet(name)
    ws.append(["姓名", "年龄"])
    for i in range(1, 2401):
        ws.append(["学生%d" % i, 7 + i % 10])
    ws.append(["坏行", "暂无"])
wb.save(sys.argv[1])
PY17GEN
rm -f "$TMP/errrows.xlsx" "$D/脏甲.sql" "$D/脏乙.sql"
run 0 "多 sheet 失败行聚合" "$TMP/17j.txt" "$TMP/dirty_multi.xlsx" --seed 5
[ -f "$TMP/errrows.xlsx" ] && ok "errrows.xlsx 已生成" || no "errrows.xlsx 已生成"
"$PY" - "$TMP/errrows.xlsx" > "$TMP/17k.txt" <<'PY17CHK'
import sys
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1], read_only=True)
names = wb.sheetnames
assert names == ["脏甲", "脏乙"], names
for n in names:
    rows = list(wb[n].iter_rows(values_only=True))
    assert rows[0][0] == "原行号", rows[0]
    assert any(r[0] == 2402 for r in rows[1:]), (n, rows[1:])
print("OK")
PY17CHK
if grep -q "^OK$" "$TMP/17k.txt"; then
  ok "errrows 按源 sheet 分工作表（脏甲/脏乙）"
else
  no "errrows 按源 sheet 分工作表（脏甲/脏乙）"
fi

# 目录批量模式：当前目录所有 .xls/.xlsx，各转全部 sheet
mkdir -p "$TMP/batch17"
cp "$D/08_multi_sheet.xlsx" "$D/01_normal.xlsx" "$TMP/batch17/"
(cd "$TMP/batch17" && "$BIN" --progress off --no-color > batch17.txt 2>&1)
if [ -f "$TMP/batch17/第一张表.sql" ] && [ -f "$TMP/batch17/第二张表.sql" ] \
   && [ -f "$TMP/batch17/学生信息.sql" ] \
   && grep -q "结果表 3 · 批量汇总" "$TMP/batch17/batch17.txt"; then
  ok "目录批量转换（2 个文件、汇总表）"
else
  no "目录批量转换（2 个文件、汇总表）"
fi


echo "================ 通过 $pass / 失败 $fail ================"
[ "$fail" = "0" ]
