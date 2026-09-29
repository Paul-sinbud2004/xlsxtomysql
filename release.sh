#!/usr/bin/env bash
# 打包发布产物：编译 → 归档 → 校验和
#
#   bash release.sh                                  # 只打包本机平台
#   bash release.sh --target x86_64-unknown-linux-musl,aarch64-unknown-linux-gnu
#   bash release.sh --dist /tmp/dist                 # 指定产物目录（默认 ./dist）
#   bash release.sh --no-build                       # 复用 dist/stage 里已有的二进制，只重新打包
#
# 产物：
#   dist/xlsxtomysql-<版本>-<平台>.tar.gz   （Windows 目标为 .zip）
#   dist/xlsxtomysql-<版本>-<平台>.sha256
#   dist/SHA256SUMS                          （所有产物的汇总校验和）
#
# 交叉编译需要先装目标平台标准库： rustup target add <triple>
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC="$HERE/xlsxtomysql.rs"
DIST="$HERE/dist"
BUILD=1
TARGETS=""

while [ $# -gt 0 ]; do
  case "$1" in
    --dist)     DIST="$2"; shift 2 ;;
    --target)   TARGETS="${TARGETS:+$TARGETS,}$2"; shift 2 ;;
    --no-build) BUILD=0; shift ;;
    --help|-h)  awk 'NR>1 && /^#/ { sub(/^# ?/, ""); print; next } NR>1 { exit }' "${BASH_SOURCE[0]}"; exit 0 ;;
    *) echo "× 未知参数 $1" >&2; exit 1 ;;
  esac
done

VERSION="$(bash "$HERE/build.sh" --version)"
HOST_TRIPLE="$(rustc -vV | awk '/^host: /{print $2}')"
[ -n "$TARGETS" ] || TARGETS="$HOST_TRIPLE"

STAGE="$DIST/stage"
mkdir -p "$DIST"
# 后面会在子 shell 里 cd 到 stage 再打包，DIST 必须是绝对路径
DIST="$(cd "$DIST" && pwd)"
STAGE="$DIST/stage"
mkdir -p "$STAGE"

# triple → 便于人读的平台名（发行包文件名用它）
platform_name() {
  case "$1" in
    x86_64-unknown-linux-gnu)     echo "linux-x86_64" ;;
    x86_64-unknown-linux-musl)    echo "linux-x86_64-musl" ;;
    i686-unknown-linux-gnu)       echo "linux-i686" ;;
    aarch64-unknown-linux-gnu)    echo "linux-aarch64" ;;
    aarch64-unknown-linux-musl)   echo "linux-aarch64-musl" ;;
    armv7-unknown-linux-gnueabihf) echo "linux-armv7" ;;
    x86_64-apple-darwin)          echo "macos-x86_64" ;;
    aarch64-apple-darwin)         echo "macos-arm64" ;;
    x86_64-pc-windows-gnu)        echo "windows-x86_64" ;;
    x86_64-pc-windows-msvc)       echo "windows-x86_64" ;;
    aarch64-pc-windows-msvc)      echo "windows-arm64" ;;
    *)                            echo "$1" ;;
  esac
}

sha256_of() {
  if command -v sha256sum >/dev/null 2>&1; then sha256sum "$1" | awk '{print $1}'
  elif command -v shasum >/dev/null 2>&1; then shasum -a 256 "$1" | awk '{print $1}'
  else echo "× 找不到 sha256sum / shasum" >&2; exit 1
  fi
}

rm -f "$DIST"/SHA256SUMS
: > "$DIST/SHA256SUMS"

IFS=',' read -r -a LIST <<< "$TARGETS"
for TRIPLE in "${LIST[@]}"; do
  [ -n "$TRIPLE" ] || continue
  NAME="$(platform_name "$TRIPLE")"
  PKG="xlsxtomysql-$VERSION-$NAME"
  ROOT="$STAGE/$PKG"
  mkdir -p "$ROOT"

  case "$TRIPLE" in
    *windows*) BIN_NAME="xlsxtomysql.exe" ;;
    *)         BIN_NAME="xlsxtomysql" ;;
  esac

  if [ "$BUILD" = 1 ]; then
    if [ "$TRIPLE" = "$HOST_TRIPLE" ]; then
      bash "$HERE/build.sh" --release --out "$ROOT/$BIN_NAME"
    else
      bash "$HERE/build.sh" --release --target "$TRIPLE" --out "$ROOT/$BIN_NAME"
    fi
  fi
  [ -f "$ROOT/$BIN_NAME" ] || { echo "× 缺少 $ROOT/$BIN_NAME（--no-build 模式下请先构建）" >&2; exit 1; }

  # 组装包内容：二进制 + 文档 + man page
  cp -f "$HERE/README.md" "$HERE/LICENSE" "$HERE/CHANGELOG.md" "$ROOT/"
  [ -f "$HERE/README.zh-CN.md" ] && cp -f "$HERE/README.zh-CN.md" "$ROOT/"
  mkdir -p "$ROOT/docs"
  [ -f "$HERE/docs/xlsxtomysql.1" ] || python3 "$HERE/tools/gen_man.py" "$SRC" >/dev/null
  [ -f "$HERE/docs/xlsxtomysql.zh.1" ] || python3 "$HERE/tools/gen_man.py" "$SRC" --lang zh >/dev/null
  cp -f "$HERE/docs/xlsxtomysql.1" "$HERE/docs/xlsxtomysql.zh.1" "$ROOT/docs/"
  cp -f "$HERE/install.sh" "$ROOT/"
  if [ -f "$HERE/install.ps1" ]; then cp -f "$HERE/install.ps1" "$ROOT/"; fi
  cat > "$ROOT/BUILD-INFO.txt" <<EOF
包名      : $PKG
版本      : $VERSION
目标平台  : $TRIPLE
构建主机  : $HOST_TRIPLE
Rust      : $(rustc -V)
构建时间  : $(date '+%Y-%m-%d %H:%M:%S %z')
说明      : 二进制零依赖；运行时需要一个带 openpyxl（读 .xls 还需 xlrd）的 Python 解释器。
            安装脚本：Unix 用 bash install.sh；Windows 用 powershell -File install.ps1
EOF

  # 打包：Windows 用 zip（对方解压更顺手），其它平台用 tar.gz
  case "$TRIPLE" in
    *windows*)
      if command -v zip >/dev/null 2>&1; then
        ( cd "$STAGE" && zip -qr "$DIST/$PKG.zip" "$PKG" )
        OUT="$DIST/$PKG.zip"
      else
        ( cd "$STAGE" && tar czf "$DIST/$PKG.tar.gz" "$PKG" )
        OUT="$DIST/$PKG.tar.gz"
        echo "  （本机没有 zip，Windows 包改成了 tar.gz）"
      fi
      ;;
    *)
      ( cd "$STAGE" && tar czf "$DIST/$PKG.tar.gz" "$PKG" )
      OUT="$DIST/$PKG.tar.gz"
      ;;
  esac

  SUM="$(sha256_of "$OUT")"
  printf '%s  %s\n' "$SUM" "$(basename "$OUT")" > "$OUT.sha256"
  printf '%s  %s\n' "$SUM" "$(basename "$OUT")" >> "$DIST/SHA256SUMS"
  echo "✓ $OUT  ($(du -h "$OUT" | cut -f1))"
done

echo
echo "== 产物 =="
ls -lh "$DIST" | grep -v '^total' | grep -v stage
echo
echo "校验方式： sha256sum -c SHA256SUMS   （macOS 用 shasum -a 256 -c SHA256SUMS）"
echo "产物目录： $DIST"
