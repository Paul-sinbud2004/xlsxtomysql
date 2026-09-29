#!/usr/bin/env bash
# 编译 xlsxtomysql（Rust 单文件版）
#
#   bash build.sh                     # 编译到 ./xlsxtomysql
#   bash build.sh /usr/local/bin/xlsxtomysql
#   bash build.sh --release           # 发布构建（opt-level=3 + LTO + 去符号表）
#   bash build.sh --release --target x86_64-unknown-linux-musl
#   bash build.sh --sync              # 把 py_helper.py 回写进 xlsxtomysql.rs 的 PY_HELPER 常量
#   bash build.sh --version           # 只打印源码里的版本号（发布脚本用）
#
# 只依赖 rustc，不需要 cargo、不需要任何 crate、不需要联网。
# 需要交叉编译时，先用 rustup 装目标平台的标准库：
#     rustup target add x86_64-pc-windows-gnu
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC="$HERE/xlsxtomysql.rs"

SYNC_PY="$(command -v python3 || true)"

RELEASE=0
TARGET=""
OUT=""

while [ $# -gt 0 ]; do
  case "$1" in
    --sync)
      [ -n "$SYNC_PY" ] || { echo "× --sync 需要 python3" >&2; exit 1; }
      "$SYNC_PY" - "$SRC" "$HERE/py_helper.py" <<'PY'
import re, sys
src_path, disk_path = sys.argv[1], sys.argv[2]
disk = open(disk_path, encoding="utf-8").read().rstrip("\n")
assert '"##' not in disk, "py_helper.py 里出现 \"## ，会截断 Rust 原始字符串"
s = open(src_path, encoding="utf-8").read()
m = re.search(r'(const PY_HELPER: &str = r##")(.*?)("##;)', s, re.S)
assert m, "未在 xlsxtomysql.rs 中找到 PY_HELPER 常量"
s = s[:m.start(2)] + disk + s[m.end(2):]
open(src_path, "w", encoding="utf-8").write(s)
print("✓ 已把 py_helper.py 回写入 xlsxtomysql.rs（%d 行）" % (disk.count("\n") + 1))
PY
      shift
      ;;
    --release) RELEASE=1; shift ;;
    --strip)   RELEASE=1; shift ;;   # 兼容旧写法
    --target)  TARGET="$2"; shift 2 ;;
    --out)     OUT="$2"; shift 2 ;;
    --version)
      [ -n "$SYNC_PY" ] || { echo "× --version 需要 python3" >&2; exit 1; }
      "$SYNC_PY" - "$SRC" <<'PY'
import re, sys
m = re.search(r'const VERSION: &str = "([^"]+)"', open(sys.argv[1], encoding="utf-8").read())
if not m:
    sys.exit("× 找不到 VERSION 常量")
print(m.group(1))
PY
      exit 0
      ;;
    --help|-h)
      awk 'NR>1 && /^#/ { sub(/^# ?/, ""); print; next } NR>1 { exit }' "${BASH_SOURCE[0]}"
      exit 0
      ;;
    -*) echo "× 未知参数 $1（用 --help 看用法）" >&2; exit 1 ;;
    *)  OUT="$1"; shift ;;
  esac
done

command -v rustc >/dev/null 2>&1 || {
  echo "× 未找到 rustc。请先安装 Rust 工具链：" >&2
  echo "    curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh" >&2
  echo "  或使用发行版包管理器：apt install rustc / dnf install rust" >&2
  exit 1
}

[ -f "$SRC" ] || { echo "× 找不到源码 $SRC" >&2; exit 1; }

echo "== Rust 版本 =="
rustc -V

# ---- 一致性检查：xlsxtomysql.rs 里内嵌的 Python 助手是否与 py_helper.py 同步 ----
if [ -f "$HERE/py_helper.py" ] && [ -n "$SYNC_PY" ]; then
  "$SYNC_PY" - "$SRC" "$HERE/py_helper.py" <<'PY' || true
import re, sys, difflib
src_path, disk_path = sys.argv[1], sys.argv[2]
src = open(src_path, encoding="utf-8").read()
m = re.search(r'const PY_HELPER: &str = r##"(.*?)"##;', src, re.S)
if not m:
    print("⚠ 未在源码中找到 PY_HELPER 常量，跳过一致性检查")
    sys.exit(0)
emb = m.group(1).rstrip("\n")
disk = open(disk_path, encoding="utf-8").read().rstrip("\n")
if emb == disk:
    print("✓ 内嵌 Python 助手与 py_helper.py 一致（%d 行）" % (emb.count("\n") + 1))
else:
    diff = list(difflib.unified_diff(disk.split("\n"), emb.split("\n"),
                                     "py_helper.py", "内嵌副本", lineterm="", n=1))
    print("⚠ 内嵌 Python 助手与 py_helper.py 不一致（差异 %d 行）：" % len(diff))
    print("\n".join(diff[:30]))
    print("  提示：改完 py_helper.py 后执行 `bash build.sh --sync` 回写。")
PY
fi

# ---- 推断输出文件名：交叉编译时带上目标平台后缀，免得几个平台的产物互相覆盖 ----
if [ -z "$OUT" ]; then
  if [ -n "$TARGET" ]; then
    NAME="xlsxtomysql-$TARGET"
    case "$TARGET" in
      *windows*) NAME="$NAME.exe" ;;
    esac
    OUT="$HERE/$NAME"
  else
    OUT="$HERE/xlsxtomysql"
  fi
fi

# ---- 编译参数 ----
FLAGS=(--edition 2021)
if [ "$RELEASE" = 1 ]; then
  FLAGS+=(-C opt-level=3 -C lto -C codegen-units=1 -C strip=symbols)
  echo "== 发布构建（-C opt-level=3 -C lto -C strip=symbols）=="
else
  FLAGS+=(-O)
fi
if [ -n "$TARGET" ]; then
  FLAGS+=(--target "$TARGET")
fi
FLAGS+=("$SRC" -o "$OUT")

echo "== 编译 =="
# 输出目录不存在就建出来（发布脚本会写到 .tmpchk/rel 之类的新目录）
mkdir -p "$(dirname "$OUT")"
rustc "${FLAGS[@]}"

echo
echo "✓ 编译完成: $OUT  ($(du -h "$OUT" | cut -f1))"
echo
echo "运行前请确认有可用的 Python 解释器（需要 openpyxl；读 .xls 还需要 xlrd）："
echo "    \"$OUT\" tests/samples/01_normal.xlsx 学生信息 students 1 2"
echo "可用 --python /path/to/python 指定解释器，或设环境变量 XLSXTOMYSQL_PYTHON。"
