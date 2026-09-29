# 更新记录 / Changelog

版本号规则：`主版本.次版本.修订号`（语义化版本）。
开发期曾用 `-rs` 后缀区分同名的 Python 实现，自首个公开发布版起统一去掉后缀。

所有值得注意的改动都会记在这里。格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## 2.3.0 —— 多工作表与批量模式

### 新增

- **所有位置参数都可省略**，按「文件名 sheet 新表名 字段行 数据行 共几行」的顺序依次填充：
  - 省略「文件名」→ 批量转换当前目录下所有 .xls / .xlsx 文件（跳过 `~$` 临时文件），
    每个文件转全部工作表，最后输出「结果表 3 · 批量汇总」。
  - 省略「sheet」→ 转换该文件的所有工作表，每个 sheet 各生成一个 `<sheet名>.sql`。
  - 省略「新表名」→ 用 sheet 名做表名（非法字符自动处理、超长截断、重名追加 `_2`/`_3`）。
  - 省略行号 → 字段名称所在行 = 1、第一个数据所在行 = 2、共几行 = 到文件尾。
- **sheet 参数支持编号与区间表达式**：`2` 表示第 2 个工作表；`[1-3]`、`[1,3,5]`、
  `[1,3-5]` 表示多个（方括号可省，中英文逗号/分号均可分隔）。
  匹配顺序：先按工作表名精确匹配，再按编号/表达式解析。
- **errrows.xlsx 支持多工作表**：多 sheet / 批量模式下，所有失败行汇总进同一个
  errrows.xlsx，按源 sheet 分工作表记录（表名沿用源 sheet 名，Excel 31 字符内去重）。
- 多 sheet 模式新增「结果表 3 · 批量汇总」，逐 sheet 列出成功/失败行数与输出文件。
- 空 sheet（无数据区/空表头）在多 sheet 模式下自动跳过并在报表中说明。

### 行为变更

- 「新表名」位置上的纯数字一律按行号解析（表名不能是纯数字）。
- 多 sheet / 批量模式下 `--out` 不可用（各 sheet 各自输出），指定新表名也不可用。
- 退出码聚合：任一 sheet 被非标准格式阻断 → 2；批量模式下某个文件无法读取 → 1。

## 2.2.0 —— 首个公开发布版

### 新增

- **中英双语界面**。默认跟随系统语言：中文环境输出中文，其它环境输出英文。
  判定顺序 `--lang` > `XLSXTOMYSQL_LANG` > `LC_ALL` > `LC_MESSAGES` > `LANG`；
  一个都没有时，Unix 按英文、Windows 读系统界面语言。
  生成的 `.sql` 注释、`--report` 报表、`errrows.xlsx` 表头都跟着同一个语言走。
- **`--man` 详细手册**（中英各一份），涵盖选项细节、类型推断规则、抽样算法、
  特殊表格处理、退出码与已知限制。`--help` 增加语言与 `--man` 说明。
- **`--version` 带上目标平台**，例如 `xlsxtomysql 2.2.0 (linux-x86_64)`。
- **安装与发布配套**：`install.sh`（Unix）、`install.ps1`（Windows）、
  `release.sh`（打包 + SHA256SUMS）、`LICENSE`、`CHANGELOG.md`、
  `Cargo.toml`（便于 `cargo build`）、`docs/xlsxtomysql.1`（man page）。
- **校验工具**：`tools/check_help_sync.py`（选项 ↔ 文档同步）、
  `tools/i18n_inventory.py` + `tools/apply_i18n.py`（文案清单与译文表覆盖率）、
  `tests/check_en_clean.py`（英文模式无中文残留）。

### 改进（跨平台）

- **Windows**：内嵌 Python 助手源码约 35 KB，超过 `CreateProcess` 约 32 K 的命令行上限，
  用 `python -c <源码>` 会直接启动失败。现在 Windows 上自动改为写入临时脚本再执行；
  Unix 继续用 `-c`，磁盘上不落地 `.py`。可用 `XLSXTOMYSQL_PYFILE=1` 在本机演练该路径。
- **Windows**：控制台自动切到 UTF-8 代码页并打开 ANSI 转义，彩色与中文不再乱码；
  解释器查找按 `PATHEXT` 补全 `.exe`，并支持 `py` 启动器与 `%LOCALAPPDATA%` 下的 venv。
- **Unix**：恢复默认 `SIGPIPE` 行为，`xlsxtomysql --man | head` 之类不会再抛 panic 与回溯。
- **时区**：生成的 SQL 头部时间改用本地时区偏移（原实现写死 +8，非东八区用户时间不对）。
- **路径**：家目录、输出路径统一用平台原生分隔符；命令行参数按有损 UTF-8 解析，
  文件名含生僻字节时不再 panic。
- **输出原子化**：`.sql` 先写 `<目标>.part`，成功后才改名到目标。中途失败不会留下
  半截 sql 冒充成功结果，也不会覆盖上一次的成果。输出目录不存在时给出明确提示。

### 修复

- `--help` / `--man` 补齐缺失选项（`--sample-min`、`--sample-max`、`--hints`、
  `--print-errors`）；修正帮助文本里 `XLSXTOMYSQL_LANG` 的拼写错误。
- 英文模式下表格省略号改用 `...`；英文句子较长，报表单元格截断宽度相应放宽。
- 字段名提示、空行判定等拼接文案改为整句模板，避免中英混排出现病句。

## 2.1.0

- 抽出 `py_helper.py` 作为可独立单测的助手源码，`build.sh --sync` 回写进 `PY_HELPER` 常量。
- 双线程分别接管助手 stdout/stderr，进度条与事件流互不干扰。

## 2.0.0

- 首个 Rust 单文件版本：零 crate、`rustc` 一条命令编译，行为与 Python 版对齐。
- 行协议事件流（`#SHEET` / `#MERGES` / `#ROW` / `#FATAL` …），退出码 0/1/2/3。
- 类型推断、抽样跳跃扫描、合并单元格补齐、非标准格式阻断、失败行导出 `errrows.xlsx`。
