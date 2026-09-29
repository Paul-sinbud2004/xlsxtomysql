# xlsxtomysql

**把 Excel（`.xlsx` / `.xls`）转成 MySQL 建表语句 + 插入语句的命令行工具。**

[![Release](https://img.shields.io/github/v/release/Paul-sinbud2004/xlsxtomysql?sort=semver)](https://github.com/Paul-sinbud2004/xlsxtomysql/releases)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![AUR](https://img.shields.io/aur/version/xlsxtomysql)](https://aur.archlinux.org/packages/xlsxtomysql)
[![Platform](https://img.shields.io/badge/platform-linux%20%7C%20macOS%20%7C%20windows-lightgrey.svg)](#跨平台)
[![Zero crates](https://img.shields.io/badge/dependencies-none-success.svg)](#架构)

**Rust 单文件实现，零 crate 依赖，一条 `rustc` 命令编译**；Excel 的读写交给一段
**内嵌在 Rust 源码常量里的 Python 助手**，运行时不会在磁盘上落地任何 `.py` 文件。

[English](README.md) · [简体中文](README.zh-CN.md)

```bash
xlsxtomysql 文件名.xlsx  sheet名  新表名  字段名称所在行  第一个数据所在行  [共几行]
```

---

## 目录

- [它解决什么问题](#它解决什么问题)
- [安装](#安装)
- [运行前提](#运行前提)
- [用法](#用法)
- [功能清单](#功能清单)
- [特殊表格的处理](#特殊表格的处理)
- [架构](#架构)
- [测试](#测试)
- [语言与本地化](#语言与本地化)
- [跨平台](#跨平台)
- [发布打包](#发布打包)
- [常见问题](#常见问题)
- [许可](#许可)

## 它解决什么问题

把一张表格塞进 MySQL，通常只有三条路：图形化导入向导（类型经常猜错）、手写
`CREATE TABLE`（类型凭感觉）、或者写个一次性脚本。`xlsxtomysql` 把这份枯燥活干扎实：

- **按数据推断表结构**，而不是只看首行——`00123` 这种写法不会被推断成整数而丢掉前导零。
- **不静默糟改脏表格**：数据区中间的空行、分隔装饰行、重复表头、合计小计行，全部
  连同**原文件行号**逐条报出，命中阻断项就停下，而不是产出一张半对半错的表。
- **大文件用抽样跳跃扫描**，三万行不必全量读一遍才能看到推断出来的字段类型。
- **转不动的行导出到 `errrows.xlsx`**，不会有哪一行被悄悄丢掉。
- **Linux / macOS / Windows 行为一致**，中英双语文案自动切换。

## 安装

### Arch Linux（AUR）

```bash
yay -S xlsxtomysql      # 或：paru -S xlsxtomysql
```

### 预编译二进制

从 [Releases](https://github.com/Paul-sinbud2004/xlsxtomysql/releases) 下载对应平台的
压缩包（内含二进制、两份 man page、文档与安装脚本）：

```bash
# Linux / macOS
tar xzf xlsxtomysql-<版本>-<平台>.tar.gz
cd xlsxtomysql-<版本>-<平台>
bash install.sh                       # 装到 ~/.local
bash install.sh --prefix /usr/local   # 装到系统目录（可能需要 sudo）
bash install.sh --no-python           # 不建专用 venv
```

```powershell
# Windows（PowerShell，在解压出来的目录里执行）
powershell -ExecutionPolicy Bypass -File install.ps1
powershell -ExecutionPolicy Bypass -File install.ps1 -Prefix "D:\tools"
```

安装脚本只做三件事：二进制装到 `<prefix>/bin`；man page 装到 `<prefix>/share/man/man1`；
建一个专用 venv（`<prefix>/share/xlsxtomysql/venv`）并装上 `openpyxl` / `xlrd`。
程序启动时**会先找「自己和同一个 prefix 下的 venv」**，所以装完直接能用，
不需要改 PATH（脚本会提示怎么改）、也不需要设任何环境变量。

卸载就是把装进去的东西删掉，脚本不碰别的任何文件：

```
<prefix>/bin/xlsxtomysql
<prefix>/share/man/man1/xlsxtomysql*.1
<prefix>/share/xlsxtomysql/
```

### 从源码编译

只依赖 `rustc`，**不需要 cargo、不需要任何 crate、不需要联网**：

```bash
git clone https://github.com/Paul-sinbud2004/xlsxtomysql.git
cd xlsxtomysql
bash build.sh --release          # opt-level=3 + LTO + 去符号表（844 KB）
# 等价于手工一条命令：
rustc --edition 2021 -O xlsxtomysql.rs -o xlsxtomysql
```

顺手也能用 cargo（`Cargo.toml` 里把 `xlsxtomysql.rs` 指为 bin 目标，无任何依赖）：

```bash
cargo build --release
```

`build.sh` 每次都会校验 `xlsxtomysql.rs` 里内嵌的 Python 助手是否与 `py_helper.py`
保持同步，不一致会打印差异并提示 `--sync`。

## 运行前提

Rust 二进制本身零依赖，但**读取 Excel 需要机器上有一个 Python 解释器**，并装上：

| 包 | 用途 | 必需性 |
|---|---|---|
| `openpyxl` | 读 `.xlsx` | 必需 |
| `xlrd` | 读 Excel 2003 `.xls` | 只处理 `.xls` 时必需 |

写 `errrows.xlsx` 不需要任何第三方库：助手用标准库 `zipfile` 直接拼 OOXML，
不建临时文件也不删除任何文件，因此在 `/tmp` 很小或进程不允许删文件的环境里也能正常工作。

解释器的查找顺序：

1. `--python /path/to/python`
2. 环境变量 `XLSXTOMYSQL_PYTHON`
3. **二进制自己旁边的 venv**：`<可执行文件目录>/../share/xlsxtomysql/venv/…`
   （发布包解压后或安装脚本装出来的专用环境，便携场景就靠这条）
4. `~/.local/share/xlsxtomysql/venv/bin/python`
5. `~/.local/bin/python3`、`/usr/bin/python3`、`/usr/local/bin/python3`、`/opt/homebrew/bin/python3`
6. `PATH` 里的 `python3` / `python`；Windows 上按 `PATHEXT` 补全 `.exe`，并支持 `py` 启动器

> **Windows 注意**：内嵌的助手源码约 35 KB，超过 `CreateProcess` 约 32 K 的命令行上限，
> `python -c <源码>` 会被系统直接拒绝。所以 Windows 上程序会自动把助手源码写到
> `%TEMP%\xlsxtomysql\` 下一个带校验和的 `.py` 再执行（同一版本只写一次）；
> Unix 继续用 `-c`，磁盘上不落地 `.py`。想在本机演练这条路径：`XLSXTOMYSQL_PYFILE=1`。

找不到或缺少 openpyxl 时会给出明确提示，可以用 `--dump-python` 把内嵌的助手源码
导出来，自己塞进任意环境里单测：

```bash
./xlsxtomysql --dump-python > py_helper.py
python3 py_helper.py probe 数据.xlsx Sheet1     # 探测文件 / 工作表 / 合并单元格
```

## 用法

```bash
# 基本用法：字段名在第 1 行，数据从第 2 行开始，直到文件尾
xlsxtomysql 学生名单.xlsx Sheet1 students 1 2

# 只取 300 行；输出到指定路径；把报表另存一份
xlsxtomysql 大表.xlsx 明细 orders 2 3 300 --out ./orders.sql --report ./orders.txt

# 先只看一眼推断出来的字段类型，不生成 sql
xlsxtomysql 名单.xlsx Sheet1 students 1 2 --scan-only

# 表格格式不标准、但确认要强行转换，并把失败行导出来
xlsxtomysql 脏表.xlsx 数据 tmp 1 2 --force --err-file ./errrows.xlsx
```

位置参数（顺序固定）：

```
xlsxtomysql 文件名.xlsx  sheet名  新表名  字段名称所在行  第一个数据所在行  [共几行]
```

| 参数 | 说明 |
|---|---|
| `文件名.xlsx` | 支持 Excel 2007+（`.xlsx`）与 Excel 2003（`.xls`） |
| `sheet名` | 工作表名；不存在时会列出该文件的所有工作表 |
| `新表名` | 表名，同时作为输出文件名 `<新表名>.sql` |
| `字段名称所在行` | 表头行号，从 1 开始 |
| `第一个数据所在行` | 第一条数据行号 |
| `共几行` | 可选，省略则直到文件尾；`0` 也表示到文件尾 |

### 一个完整例子

仓库里的 `examples/students.xlsx` 是一张小表，特意放了文本日期列、手机号列和百分比列。
执行：

```bash
xlsxtomysql examples/students.xlsx 学生信息 students 1 2 \
  --out examples/students.sql --report examples/students_report.txt
```

得到 `examples/students.sql`（已随仓库提交，不装也能先看效果）：

```sql
CREATE TABLE IF NOT EXISTS `students` (
  `姓名` varchar(16),
  `班级` varchar(16),
  `出生日期` date,          -- 表里是文本，识别成日期
  `身高cm` decimal(4,1),
  `手机号` varchar(32),     -- 像标识字段，按文本存，前导零不会丢
  `午餐费` decimal(4,1),
  `出勤率` decimal(4,3),    -- 单元格是百分比格式，存真实比值
  `备注` varchar(16)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

INSERT INTO `students` (`姓名`,`班级`,`出生日期`,...) VALUES
  ('张小明','三年二班','2017-03-05',128.5,'13800138000',12.5,0.975,'过敏：花生'),
  ...;
```

`examples/students_report.txt` 是同一趟跑出来的控制台报表。

选项：

| 选项 | 说明 |
|---|---|
| `--out PATH` | 输出 sql 路径（默认源文件同目录 `<新表名>.sql`） |
| `--err-file PATH` | 失败行文件（默认源文件同目录 `errrows.xlsx`） |
| `--report PATH` | 把控制台报表另存一份纯文本 |
| `--force` | 发现非标准格式时仍然继续转换 |
| `--scan-only` | 只扫描并输出类型预览，不生成 sql |
| `--no-hints` / `--hints` | 关闭 / 重新打开字段名语义推断（手机号走文本、金额走 decimal 等） |
| `--full-scan` | 强制全量扫描（大文件会很慢） |
| `--sample-ratio F` | 抽样比例，默认 `0.10` |
| `--sample-threshold N` | 超过多少行启用抽样跳跃扫描，默认 `2000` |
| `--sample-min N` / `--sample-max N` | 抽样行数下限/上限，默认 `1000` / `20000` |
| `--seed N` | 抽样起点随机种子，指定后每次抽样位置一致，便于复现 |
| `--batch-size N` | 每条 `INSERT` 的行数，默认 `200` |
| `--empty-as-null M` | `auto` / `always` / `never`，默认 `auto` |
| `--varchar-max N` | `varchar` 长度上限，默认 `1000` |
| `--text-max N` | `text` 上限，超过改用 `longtext`，默认 `16000` |
| `--max-ident N` | 字段名最大长度，默认 `64` |
| `--merge-scan M` | `auto` / `on` / `off`，默认 `auto` |
| `--merge-scan-limit N` | 合并预扫描的最大文件 MB，默认 `64` |
| `--no-fill-down` | 不自动补齐数据区的合并单元格 |
| `--add-id NAME` | 追加一个自增主键列 |
| `--primary-key COL` | 指定主键列 |
| `--drop-table` | 生成 `DROP TABLE IF EXISTS` |
| `--insert-ignore` | 生成 `INSERT IGNORE` |
| `--table-comment S` | 表注释 |
| `--not-null` | 所有字段加 `NOT NULL` |
| `--progress off` | 关闭进度条 |
| `--no-color` | 关闭彩色输出 |
| `--python PATH` | 指定读取 Excel 用的 Python 解释器 |
| `--print-errors N` | 控制台最多显示多少条失败行，默认 `20` |
| `--dump-python` | 打印内嵌的 Python 助手源码后退出 |
| `--lang zh\|en` | 强制输出语言；默认按系统环境自动判断 |
| `--man` | 详细手册（选项细节、类型规则、退出码、限制），中英各一份 |
| `-h` / `--help`、`-V` / `--version` | 帮助与版本（`--version` 会带上目标平台） |

### 退出码

| 码 | 含义 |
|---|---|
| `0` | 成功 |
| `1` | 运行期错误（文件损坏、不是 Excel、识别不出格式……） |
| `2` | 检测到非标准表格格式，已中止（未生成 sql） |
| `3` | 参数用法错误（文件不存在、行号颠倒、工作表/主键名不存在……） |

## 功能清单

1. **智能类型推断** — 逐列累积统计而不是只看首行：字符长度分布、整数位数、小数位数、
   日期时间值，再结合字段名语义推断。长文本按实际长度自升至 `varchar(n)` → `text` →
   `mediumtext`，超长内容直接 `longtext`；整数按实际范围选 `tinyint` / `smallint` /
   `int` / `bigint`；`手机 / 身份证 / 工号 / 银行卡` 一类标识字段强制按文本存，
   `金额 / 费用 / 价格` 走 `decimal`。
2. **非标准格式拦截** — 扫描阶段检测五类问题，逐条给出**原文件行号（必要时含列号）**：
   数据区中间空行、分隔线/装饰行、重复表头、合计小计行（警告级）、数据超出字段名
   列范围。命中阻断项时**不生成 sql**，退出码 `2`，提示处理源文件后重跑；加 `--force`
   可强行继续。
3. **大文件抽样跳跃扫描** — 超过阈值（默认 2000 行）自动切换为定步长跳跃抽样。
   步长取「最接近 `1/抽样比例` 的**质数**」（默认 10% → 每 11 行取 1 行），开头
   1000 行无条件参与判定，读取行数也有上限，`--seed N` 可复现抽样起点。用质数步长
   是为了避开与数据自身周期的共振：步长取 10 而数据恰好以 50 行为周期时，被抽到的
   行号模 50 永远只落在同样的 5 个余数上，会系统性漏掉一批取值（例如金额列只在别的
   余数上出现两位小数，抽样就把该列判成一位小数，转换阶段大批行失败）。报表中明确
   标注扫描方式与精度。进度条写 stderr（TTY 下是动态刷新条，含已用/剩余时间；
   重定向时降级为百分比日志行），不污染可重定向的报表。
4. **输出命名与编码** — 默认在源文件同目录生成 `<新表名>.sql`，全文件 `utf8mb4`
   （`DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci`）。
5. **字段名清洗** — 非法字符转下划线、超长截断、空字段名补 `col_d`、重名自动加 `_2`
   后缀；每处改动都在报表里说明改成了什么。
6. **报表与失败行** — 结果表 1 给出字段名行 / 首数据行 / 总行数 / 成功 / 失败 / 耗时，
   失败明细按**原文件行号 + 原因**列出并按原因归类；失败行连同原因导出为
   `errrows.xlsx`（列：`原行号` + 原字段名 + `失败原因`）。结果表 2 是
   「新字段名 ↔ 类型 ↔ 推断依据」对照表。
7. **其它边界** — Excel 2003 `.xls`、多工作表、合并表头自动向右填充、公式列提示、
   `.xlsb` 明确拒绝、空单元格三种策略、`--scan-only`、行数越界自动截断、行号颠倒 /
   负数 / 文件损坏 / 伪装成 xlsx 的文本文件等均有明确报错。

## 特殊表格的处理

| 情况 | 处理方式 |
| --- | --- |
| 合并单元格表头（`A1:D1`） | 字段名向右填充，报表里说明填充了几处 |
| 字段名行落在纵向合并区内（两行表头 `D1:D2`） | 回读合并区左上角那一行，并逐条报出合并区位置 |
| 数据区纵向合并（一个学生占 3 行） | 按左上角内容补齐整块，`--no-fill-down` 可关闭 |
| 文本日期 `20240101` / `2024-01-01T08:30:00` / `2024/1/1 8:30 AM` | 字段名含"日期/时间"时尝试识别 |
| 超过 24 小时的时长 `25:30:00`、`100:00:00` | 识别为 `time`（MySQL 上限 838:59:59） |
| 带时区偏移 `2024-01-01T08:30:00+08:00` | 按 `datetime` 入库，偏移量丢弃并明确提示 |
| 隐藏行 / 隐藏列 | 内容照常导出，并提示列名 |
| 百分比格式（显示 12.5%，实存 0.125） | 入库 0.125，并提示共有多少个 |
| WPS `DISPIMG` / Excel 单元格图片 | 图片写不进 SQL，对应单元格为空值并提示 |
| 浮动图片 / 批注 / 超链接 / 图表 | 不是单元格值，不参与导出（不影响其余列） |
| 有歧义的写法（`12/31/2024`、`€1,234.56`、文本 `12.5%`） | 一律按文本保存，不做猜测 |

"数据区中间的空行"与"末尾空行"要分开看：只有后面**又出现数据行**时，前面的空行
才算"中间空行"并阻断（退出码 2）；一路空到结尾的按末尾空行忽略——Excel 导出的表尾
常常拖着几百行空白，不能一律拦下。

## 架构

```
xlsxtomysql.rs ─┬─ Rust 主体
                │    参数解析 · 表格格式检查 · 类型推断 · SQL 生成
                │    报表渲染 · 进度条 · 失败行汇总
                │
                └─ const PY_HELPER: &str = r##"…Python 助手…"##
                     运行期以 `python3 -c <那段源码> <mode> <args…>` 调用
                     ├── probe   → 文件格式 / 工作表列表 / 尺寸 / 合并单元格 / 公式
                     ├── dump    → 流式吐出表头与数据行
                     ├── list    → 只列工作表名
                     └── errrows → 把失败行写成 errrows.xlsx（zipfile 直写，不依赖 openpyxl）
```

两个进程之间用**制表符分隔的行协议**通信：

- Python 的 **stdout** 是数据通道：`#KIND` / `#SHEET` / `#SHEETS` / `#MERGES` /
  `#FORMULA` / `#HEADER` / `#ROW` / `#END` / `#OK`，出错则是
  `#FATAL<TAB>退出码<TAB>消息`（退出码沿用宿主约定：1 运行期错误、3 用法错误），
  Rust 拿到什么码就往上抛什么码。
- Python 的 **stderr** 是进度与日志通道：`#P\t<当前>\t<总数>` 驱动进度条，其他行
  原样转发给用户。
- Rust 用**两个线程**分别接管这两条管道，所以进度条刷新不会打乱数据解析。
- 单元格编码为 `<kind>:<转义后的值>`，kind 是单字符（`e` 空 / `i` 整数 / `d` 小数 /
  `f` 科学计数 / `b` 布尔 / `s` 文本 / `D` 日期 / `T` 时间 / `M` 日期时间 /
  `x` Excel 错误值）。类型判定发生在 Python 侧读单元格的时候，
  Rust 侧只负责统计与推断。

这样做的好处是：**编译产物是单个自包含的二进制**，不依赖 cargo、不依赖任何 crate；
而 Excel 解析这种「格式千奇百怪、生态成熟库都在 Python」的部分，直接复用
`openpyxl` / `xlrd`，不用为了读 `xlsx` 手写一遍 OOXML + 共享字符串表 + 格式化
数字 + 1900/1904 日期基准的解析器。

## 测试

```bash
bash build.sh                     # 编译（零警告）
bash tests/run_tests.sh           # 功能自测，112 项：退出码 / SQL 内容 / 报表文案 / 语言 / 帮助
bash tests/compare_with_python.sh # 与 Python 实现逐字节对照：19 组常规用例
python3 tools/apply_i18n.py --check     # 译文覆盖率：源码里的中文文案是否都有英文
python3 tools/check_help_sync.py        # 代码支持的每个选项是否都写进了 --help / --man
python3 tests/check_en_clean.py ./xlsxtomysql   # 实跑样例，确认英文模式无中文残留
python3 tools/gen_man.py xlsxtomysql.rs [--lang zh]   # 从源码常量重新生成 man page
```

`tests/check_errrows.py` 单独校验导出的 `errrows.xlsx` 表头与内容。

## 语言与本地化

**默认跟随系统语言**：中文环境输出中文，其它环境输出英文。

判定顺序（先命中先用）：

| 顺序 | 来源 | 说明 |
|---|---|---|
| 1 | `--lang zh` / `--lang en` | 命令行强制 |
| 2 | `XLSXTOMYSQL_LANG` | 环境变量强制（适合写进脚本/CI） |
| 3 | `LC_ALL` → `LC_MESSAGES` → `LANG` | 以 `zh` 开头判为中文 |
| 4 | 平台默认 | Unix 落到英文（等同 C locale）；Windows 读系统界面语言，主语言为中文则中文 |

**跟着语言变的，不只是控制台**：

* `--help` / `--man` 各备中英两份；
* 生成的 `.sql` 头部与尾部注释；
* `--report` 存下来的那份纯文本报表；
* `errrows.xlsx` 的表头（`原行号` / `失败原因` ↔ `Source row` / `Failure reason`）；
* 表格里的省略号（中文 `…`、英文 `...`）、列表分隔符（`、` ↔ `, `）。

**为什么是准确翻译而不是硬编码两套输出**：所有界面文案收在源码里的 `CATALOG` 常量，
以中文原文为键；输出边界（`Style::wrap`、`Out::line`、报表单元格、错误构造、
助手的事件与异常消息）统一过一层翻译。这样加新文案时漏翻能被工具查出来：

* `tools/i18n_inventory.py` 抽源码里全部面向用户的中文文案；
* `tools/apply_i18n.py --check` 核对覆盖率与占位符个数；
* `tests/check_en_clean.py` 真跑一遍所有样例，扫 stdout / stderr / `.sql` / `--report`
  里有没有中文残留（这是最后一道闸）。

想加一门新语言，改 `CATALOG` 与语言判定即可，不需要碰业务代码。

`man xlsxtomysql` 装的 man page 与 `--man` 同源，由 `tools/gen_man.py` 从同一份常量生成，
不会出现「手册改了 man page 忘了改」。

## 跨平台

| | Linux | macOS | Windows |
|---|---|---|---|
| 二进制 | ✅ | ✅ | ✅ |
| 控制台 UTF-8 / 彩色 | ✅ 原生 | ✅ 原生 | 启动时切 UTF-8 代码页并打开 ANSI 转义 |
| 解释器查找 | `python3` / `python` / venv | Homebrew 路径也查 | `python.exe`（按 `PATHEXT` 补全）、`py` 启动器、`%LOCALAPPDATA%` |
| 家目录 | `$HOME` | `$HOME` | `$USERPROFILE` / `%LOCALAPPDATA%` |
| 路径拼接 | `/` | `/` | `\`（统一走 `PathBuf`） |
| 助手启动 | `python -c <源码>`（不落盘） | 同 Linux | 落临时 `.py` 再执行（命令行长度限制，见上文） |
| 本地时区 | `localtime_r` | 同 Linux | `GetLocalTime` 与 `GetSystemTime` 相减 |
| 管道截断 | 恢复 `SIGPIPE`，`\| head` 安静退出 | 同 Linux | 不适用 |

**已经实跑验证的**：Linux x86_64（编译零警告、全部自测与对照通过）。
**只做了代码分支与逻辑审查、没在真机跑过的**：Windows 与 macOS 特有分支
（`GetUserDefaultUILanguage`、控制台代码页、`PATHEXT` 查找、`py` 启动器）。
Windows 上那条「助手落临时文件」的路径可以用 `XLSXTOMYSQL_PYFILE=1` 在 Linux 上演练，
自动化测试里已经包含这一项。

交叉编译（产物命名自动带平台后缀）：

```bash
rustup target add x86_64-pc-windows-gnu aarch64-unknown-linux-gnu
bash build.sh --release --target x86_64-pc-windows-gnu
bash release.sh --target x86_64-unknown-linux-gnu,aarch64-unknown-linux-gnu,x86_64-pc-windows-gnu
```

## 发布打包

```bash
bash release.sh                       # 打包本机平台
bash release.sh --target <triple>,... # 多平台
bash release.sh --dist /tmp/dist      # 指定产物目录（默认 ./dist）
bash release.sh --no-build            # 复用 dist/stage 里已编译的二进制，只重新打包
```

产物：

```
dist/xlsxtomysql-<版本>-<平台>.tar.gz     （Windows 目标为 .zip）
dist/xlsxtomysql-<版本>-<平台>.sha256
dist/SHA256SUMS                           所有产物的汇总校验和
```

包内包含二进制、`README.md`、`LICENSE`、`CHANGELOG.md`、两份 man page、
两个安装脚本，以及一份 `BUILD-INFO.txt`（版本、目标平台、构建主机、Rust 版本、构建时间）。

**发布前检查清单**（全部通过再打标签）：

```bash
bash build.sh --release                     # 1. 发布构建，零警告
bash tests/run_tests.sh                     # 2. 功能自测全绿（112 项）
bash tests/compare_with_python.sh           # 3. 与 Python 实现逐字节等价（19 组用例）
python3 tools/apply_i18n.py --check         # 4. 译文覆盖率
python3 tools/check_help_sync.py            # 5. 帮助/手册与代码同步
python3 tests/check_en_clean.py ./xlsxtomysql  # 6. 英文模式无中文残留
python3 tools/gen_man.py xlsxtomysql.rs && python3 tools/gen_man.py xlsxtomysql.rs --lang zh
                                            # 7. man page 与手册常量同步
bash release.sh                             # 8. 出包 + 校验和
```

版本号出现在三处，改版本时一起改：`xlsxtomysql.rs` 的 `VERSION`、`Cargo.toml` 的 `version`、
`CHANGELOG.md` 的条目。`build.sh --version` 读的是源码里的那个，`release.sh` 用它给产物命名。
改了 `MAN_EN` / `MAN_ZH` 里的版本行，记得重跑第 7 步，否则第 2 步会失败。

## 常见问题

**Q：一定要装 Python 吗？**
读 Excel 这一步需要。Rust 二进制本身零依赖，但 OOXML / BIFF 的解析没有内置实现。
如果你需要彻底去掉 Python，可以引入 `calamine` crate 换成纯 Rust 读取——代价是放弃
「零 crate、一条 rustc 编译」这个特性（`build.sh` 里改一行即可）。

**Q：`.xls` 报错说读不了？**
需要一个装了 `xlrd` 的解释器才行。可以用 `install.sh` 建好虚拟环境，
或用 `--python` 指向任意带 `xlrd` 的解释器。

**Q：生成的表名/字段名带反引号吗？**
带。所有标识符都用反引号包裹，内部反引号会转义为双反引号，避免关键字冲突。

**Q：为什么进度条有时候不显示动态条？**
输出被重定向到文件或管道时自动降级为百分比日志行，避免刷屏。终端里跑就是动态条。

**Q：会不会丢行？**
不会。转不动的行会被计数、按原因列出，并导出到 `errrows.xlsx`。

## 许可

MIT License，见 `LICENSE`。内嵌的 Python 助手与 Rust 主体同一许可。

## 关键词

Excel 转 MySQL · xlsx 转 SQL · xls 转 MySQL · Excel 导入数据库 · 表格转数据库 ·
生成建表语句 · 生成 INSERT 语句 · xlsx2mysql · 命令行工具 · 单文件二进制 ·
跨平台 · Linux · macOS · Windows · Arch Linux · AUR
