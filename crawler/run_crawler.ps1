$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

$script = Join-Path $PSScriptRoot "cumt_crawler.py"
python $script --max-pages 5 --delay 0.8
