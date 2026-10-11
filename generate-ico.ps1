Add-Type -AssemblyName System.Drawing

$pngPath = Join-Path $PSScriptRoot "sysgen_logo.png"
$icoPath = Join-Path $PSScriptRoot "app_icon.ico"

$src = [System.Drawing.Bitmap]::FromFile($pngPath)
$bmp = New-Object System.Drawing.Bitmap 128, 128
$gfx = [System.Drawing.Graphics]::FromImage($bmp)
$gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$gfx.DrawImage($src, 0, 0, 128, 128)
$gfx.Dispose()

$hIcon = $bmp.GetHicon()
$icon = [System.Drawing.Icon]::FromHandle($hIcon)
$fs = [System.IO.File]::OpenWrite($icoPath)
$icon.Save($fs)
$fs.Close()
$bmp.Dispose()
$src.Dispose()

Write-Host "ICO Generated: $icoPath"
