#!/usr/bin/env bash
# xlsxtomysql 安装脚本（Linux / macOS）
#
#   bash install.sh                    # 装到 ~/.local（二进制 + man page + Python 依赖）
#   bash install.sh --prefix /usr/local
#   bash install.sh --no-python        # 不建 venv（你已有带 openpyxl 的解释器）
#   bash install.sh --no-man           # 不装 man page
#   bash install.sh --build            # 强制从源码重新编译（需要 rustc）
#
# 脚本做的事：
#   1. 找二进制：优先用同目录下的 xlsxtomysql；没有就（--build 时）用 rustc 编译一份；
#   2. 装到 <prefix>/bin/xlsxtomysql，man page 装到 <prefix>/share/man/man1/；
#   3. 建一个专用 venv（<prefix>/share/xlsxtomysql/venv）并装上 openpyxl / xlrd，
#      这样程序运行时能自己找到解释器，不用你改 PATH 或设环境变量。
#
# 卸载：删掉 <prefix>/bin/xlsxtomysql、<prefix>/share/man/man1/xlsxtomysql*.1、
#       <prefix>/share/xlsxtomysql 即可。脚本不动别的任何东西。
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PREFIX="${HOME}/.local"
DO_PYTHON=1
DO_MAN=1
DO_BUILD=0

while [ $# -gt 0 ]; do
  case "$1" in
    --prefix)    PREFIX="$2"; shift 2 ;;
    --no-python) DO_PYTHON=0; shift ;;
    --no-man)    DO_MAN=0; shift ;;
    --build)     DO_BUILD=1; shift ;;
    --help|-h)   awk 'NR>1 && /^#/ { sub(/^# ?/, ""); print; next } NR>1 { exit }' "${BASH_SOURCE[0]}"; exit 0 ;;
    *) echo "× 未知参数 $1（用 --help 看用法）" >&2; exit 1 ;;
  esac
done

say() { printf '%s\n' "$*"; }
die() { printf '× %s\n' "$*" >&2; exit 1; }

BIN_DIR="$PREFIX/bin"
MAN_DIR="$PREFIX/share/man/man1"
VENV_DIR="$PREFIX/share/xlsxtomysql/venv"

# ---------------------------------------------------------------- 1. 找二进制
SRC_BIN=""
for cand in "$HERE/xlsxtomysql" "$HERE/xlsxtomysql.exe" "$HERE/../xlsxtomysql"; do
  if [ -f "$cand" ] && [ -x "$cand" ]; then SRC_BIN="$cand"; break; fi
done

if [ "$DO_BUILD" = 1 ] || [ -z "$SRC_BIN" ]; then
  if [ -f "$HERE/xlsxtomysql.rs" ]; then
    command -v rustc >/dev/null 2>&1 || die "没有现成二进制，也没找到 rustc。
     请安装 Rust（curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh），
     或从发布页下载对应平台的压缩包后再运行本脚本。"
    say "== 从源码编译（rustc $(rustc -V | awk '{print $2}')）=="
    ( cd "$HERE" && bash build.sh --release --out "$HERE/xlsxtomysql" )
    SRC_BIN="$HERE/xlsxtomysql"
  else
    die "本目录里找不到可执行文件 xlsxtomysql，也没有源码 xlsxtomysql.rs"
  fi
fi
say "✓ 二进制: $SRC_BIN"

# ---------------------------------------------------------------- 2. 装二进制
mkdir -p "$BIN_DIR"
install -m 0755 "$SRC_BIN" "$BIN_DIR/xlsxtomysql"
say "✓ 已安装: $BIN_DIR/xlsxtomysql  ($("$BIN_DIR/xlsxtomysql" --version))"

# ---------------------------------------------------------------- 3. man page
if [ "$DO_MAN" = 1 ]; then
  MANS=""
  for m in "$HERE/docs/xlsxtomysql.1" "$HERE/docs/xlsxtomysql.zh.1" "$HERE/xlsxtomysql.1"; do
    [ -f "$m" ] && MANS="$MANS $m"
  done
  if [ -z "$MANS" ] && [ -f "$HERE/tools/gen_man.py" ] && command -v python3 >/dev/null 2>&1; then
    python3 "$HERE/tools/gen_man.py" "$HERE/xlsxtomysql.rs" >/dev/null
    python3 "$HERE/tools/gen_man.py" "$HERE/xlsxtomysql.rs" --lang zh >/dev/null
    MANS=" $HERE/docs/xlsxtomysql.1 $HERE/docs/xlsxtomysql.zh.1"
  fi
  if [ -n "$MANS" ]; then
    mkdir -p "$MAN_DIR"
    for m in $MANS; do install -m 0644 "$m" "$MAN_DIR/$(basename "$m")"; done
    say "✓ man page: $MAN_DIR/（man xlsxtomysql）"
  else
    say "· 没找到 man page，跳过（可用 ./xlsxtomysql --man 看内置手册）"
  fi
fi

# ---------------------------------------------------------------- 4. Python 依赖
if [ "$DO_PYTHON" = 1 ]; then
  if [ -x "$VENV_DIR/bin/python" ]; then
    say "· 专用 venv 已存在，检查依赖"
  else
    say "== 创建专用 Python 环境: $VENV_DIR =="
    mkdir -p "$(dirname "$VENV_DIR")"
    if command -v uv >/dev/null 2>&1; then
      # 系统 Python 在只读目录、或发行版禁止 pip 时，uv 更省事
      uv venv --system-site-packages "$VENV_DIR" >/dev/null
    elif command -v python3 >/dev/null 2>&1; then
      python3 -m venv --system-site-packages "$VENV_DIR" >/dev/null
    else
      say "⚠ 找不到 python3，跳过。程序运行需要一个带 openpyxl 的解释器。"
      DO_PYTHON=0
    fi
  fi

  if [ "$DO_PYTHON" = 1 ]; then
    PY="$VENV_DIR/bin/python"
    if ! "$PY" -c 'import openpyxl' >/dev/null 2>&1; then
      say "== 安装 openpyxl / xlrd =="
      if command -v uv >/dev/null 2>&1; then
        uv pip install --python "$PY" openpyxl xlrd || \
          say "⚠ 依赖安装失败（可稍后手动装：$PY -m pip install openpyxl xlrd）"
      else
        "$PY" -m pip install --quiet --upgrade pip >/dev/null 2>&1 || true
        "$PY" -m pip install openpyxl xlrd || \
          say "⚠ 依赖安装失败（可稍后手动装：$PY -m pip install openpyxl xlrd）"
      fi
    fi
    if "$PY" -c 'import openpyxl' >/dev/null 2>&1; then
      if "$PY" -c 'import xlrd' >/dev/null 2>&1; then
        say "✓ Python 依赖就绪（openpyxl + xlrd）"
      else
        say "✓ Python 依赖就绪（openpyxl；缺 xlrd，读 .xls 时需要）"
      fi
    else
      say "⚠ 没装上 openpyxl，运行时请用 --python 指定一个装了它的解释器"
    fi
  fi
fi

# ---------------------------------------------------------------- 5. 收尾提示
say ""
case ":$PATH:" in
  *":$BIN_DIR:"*) say "可以直接用了：  xlsxtomysql --help" ;;
  *) say "⚠ $BIN_DIR 不在 PATH 里，先把它加进去："
     say "    # zsh"
     say "    echo 'export PATH=\"$BIN_DIR:\$PATH\"' >> ~/.zshrc && exec zsh"
     say "    # bash"
     say "    echo 'export PATH=\"$BIN_DIR:\$PATH\"' >> ~/.bashrc && exec bash"
     say "  临时用也可以直接写全路径： $BIN_DIR/xlsxtomysql --help" ;;
esac
say ""
say "试试："
say "    xlsxtomysql 表格.xlsx Sheet1 新表名 1 2"
say "    xlsxtomysql --man        # 详细手册（中英随系统语言）"
