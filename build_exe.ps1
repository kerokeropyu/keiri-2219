# keiri-2219.exe を dist フォルダに作成する。
# 実行方法: powershell -ExecutionPolicy Bypass -File .\build_exe.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

uv sync
if ($LASTEXITCODE -ne 0) { throw "uv sync に失敗しました。" }

uv run pyinstaller --onefile --console --noconfirm --name keiri-2219 --paths src src/keiri_2219/cli.py
if ($LASTEXITCODE -ne 0) { throw "exeの作成に失敗しました。" }

Write-Host "作成しました: $PSScriptRoot\dist\keiri-2219.exe"
