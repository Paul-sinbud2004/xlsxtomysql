<#
xlsxtomysql 安装脚本（Windows / PowerShell）

用法（在解压出来的目录里执行）：
    powershell -ExecutionPolicy Bypass -File install.ps1
    powershell -ExecutionPolicy Bypass -File install.ps1 -Prefix "D:\tools"
    powershell -ExecutionPolicy Bypass -File install.ps1 -NoPython -NoMan

脚本做的事：
  1. 把同目录下的 xlsxtomysql.exe 复制到 <Prefix>\bin；
  2. 把 <Prefix>\bin 加进当前用户的 PATH（只改用户级环境变量，不动系统级）；
  3. 建一个专用 venv（<Prefix>\share\xlsxtomysql\venv）并装上 openpyxl / xlrd，
     程序运行时会自己找到它，不用你设环境变量。

卸载：删掉 <Prefix>\bin\xlsxtomysql.exe 与 <Prefix>\share\xlsxtomysql 目录，
     再把 <Prefix>\bin 从用户 PATH 里去掉即可。
#>
[CmdletBinding()]
param(
    [string]$Prefix = "$env:LOCALAPPDATA\Programs\xlsxtomysql",
    [switch]$NoPython,
    [switch]$NoMan
)

$ErrorActionPreference = "Stop"
function Say($m) { Write-Host $m }
function Die($m) { Write-Host "× $m" -ForegroundColor Red; exit 1 }

$Here   = Split-Path -Parent $MyInvocation.MyCommand.Path
$BinDir = Join-Path $Prefix "bin"
$VenvDir = Join-Path $Prefix "share\xlsxtomysql\venv"

# ---------------------------------------------------------------- 1. 找 exe
$SrcExe = $null
foreach ($c in @((Join-Path $Here "xlsxtomysql.exe"), (Join-Path $Here "xlsxtomysql"))) {
    if (Test-Path $c) { $SrcExe = $c; break }
}
if (-not $SrcExe) { Die "本目录里找不到 xlsxtomysql.exe。请先解压发布包，或在有 rustc 的机器上执行 bash build.sh --release。" }
Say "✓ 二进制: $SrcExe"

# ---------------------------------------------------------------- 2. 装二进制
New-Item -ItemType Directory -Force -Path $BinDir | Out-Null
$DestExe = Join-Path $BinDir "xlsxtomysql.exe"
Copy-Item -Force $SrcExe $DestExe
$ver = & $DestExe --version
Say "✓ 已安装: $DestExe  ($ver)"

# ---------------------------------------------------------------- 3. man page（有就放过去，Windows 下主要看 --man）
if (-not $NoMan) {
    $ManSrc = Join-Path $Here "docs"
    if (Test-Path $ManSrc) {
        $ManDst = Join-Path $Prefix "share\man\man1"
        New-Item -ItemType Directory -Force -Path $ManDst | Out-Null
        Copy-Item -Force (Join-Path $ManSrc "xlsxtomysql*.1") $ManDst -ErrorAction SilentlyContinue
        Say "✓ man page: $ManDst（Windows 下用 ./xlsxtomysql --man 更方便）"
    }
}

# ---------------------------------------------------------------- 4. PATH
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($null -eq $userPath) { $userPath = "" }
$parts = $userPath -split ';' | Where-Object { $_ -ne "" }
if ($parts -notcontains $BinDir) {
    $newPath = (($parts + $BinDir) -join ';')
    [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
    $env:Path = "$env:Path;$BinDir"
    Say "✓ 已把 $BinDir 加入用户 PATH（重开一个终端后生效）"
} else {
    Say "· $BinDir 已在 PATH 中"
}

# ---------------------------------------------------------------- 5. Python 依赖
if (-not $NoPython) {
    $py = $null
    foreach ($cand in @("py -3", "python", "python3")) {
        $exe = $cand.Split(' ')[0]
        if (Get-Command $exe -ErrorAction SilentlyContinue) { $py = $cand; break }
    }
    if (-not $py) {
        Say "⚠ 没有找到 Python。程序读取 Excel 需要一个带 openpyxl 的 Python 解释器。"
        Say "  安装建议：winget install Python.Python.3.12   （或到 python.org 下载安装，勾选 Add to PATH）"
    } else {
        $VenvPy = Join-Path $VenvDir "Scripts\python.exe"
        if (-not (Test-Path $VenvPy)) {
            Say "== 创建专用 Python 环境: $VenvDir =="
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $VenvDir) | Out-Null
            Invoke-Expression "$py -m venv --system-site-packages `"$VenvDir`""
        }
        if (Test-Path $VenvPy) {
            $ok = & $VenvPy -c "import openpyxl" 2>$null
            if ($LASTEXITCODE -ne 0) {
                Say "== 安装 openpyxl / xlrd =="
                & $VenvPy -m pip install --quiet --upgrade pip | Out-Null
                & $VenvPy -m pip install openpyxl xlrd
            }
            if ((& $VenvPy -c "import openpyxl" 2>$null; $LASTEXITCODE) -eq 0) {
                Say "✓ Python 依赖就绪"
            } else {
                Say "⚠ openpyxl 没装上，运行时请用 --python 指定解释器"
            }
        }
    }
}

Say ""
Say "试试（新开一个终端）："
Say "    xlsxtomysql 表格.xlsx Sheet1 新表名 1 2"
Say "    xlsxtomysql --help / --man"
