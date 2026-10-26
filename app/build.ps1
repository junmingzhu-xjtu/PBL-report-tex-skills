# -*- coding: utf-8 -*-
# 一键打包：先做图标，再出单文件版 + 免安装文件夹版，最后复制到 ..\软件发布\
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
$env:PYTHONIOENCODING = "utf-8"

Write-Host "[1/4] 生成图标" -ForegroundColor Cyan
python tools\make_icon.py

Write-Host "[2/4] 打包单文件版 dist\PBL大报告生成器.exe" -ForegroundColor Cyan
python -m PyInstaller --noconfirm --clean --onefile --windowed --icon "app.ico" `
  --name "PBL大报告生成器" --add-data "app.ico;." `
  --hidden-import pbltex.core --hidden-import pbltex.toolchain --paths "." app.py

Write-Host "[3/4] 打包免安装文件夹版 dist_app\" -ForegroundColor Cyan
python -m PyInstaller --noconfirm --clean --onedir --windowed --icon "app.ico" `
  --name "PBL大报告生成器" --add-data "app.ico;." `
  --hidden-import pbltex.core --hidden-import pbltex.toolchain --paths "." `
  --distpath dist_app --workpath build_app app.py

Write-Host "[4/4] 复制到 软件发布\" -ForegroundColor Cyan
$out = Join-Path (Split-Path $PSScriptRoot -Parent) "软件发布"
New-Item -ItemType Directory -Force $out | Out-Null
Copy-Item "dist\PBL大报告生成器.exe" (Join-Path $out "PBL大报告生成器.exe") -Force
$zip = Join-Path $out "PBL大报告生成器-免安装文件夹版.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path "dist_app\PBL大报告生成器" -DestinationPath $zip -CompressionLevel Optimal

Write-Host "完成：" -ForegroundColor Green
Get-ChildItem $out | ForEach-Object { "  {0}  {1:N1} MB" -f $_.Name, ($_.Length / 1MB) }

# 自检：界面能否启动（不弹窗模式下只写日志）
Write-Host "自检 --selftest …" -ForegroundColor Cyan
& "dist_app\PBL大报告生成器\PBL大报告生成器.exe" --selftest
Write-Host "自检退出码 $LASTEXITCODE（0 = 界面可启动）" -ForegroundColor Green
