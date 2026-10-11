$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "Sysgen RPO Loans.lnk"
$targetPath = "c:\Users\teyo\Downloads\Sysgen RPO loans compute\Sysgen-RPO-Loans.exe"
$workingDir = "c:\Users\teyo\Downloads\Sysgen RPO loans compute"
$iconPath = "c:\Users\teyo\Downloads\Sysgen RPO loans compute\app_icon.ico"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $targetPath
$Shortcut.WorkingDirectory = $workingDir
$Shortcut.IconLocation = "$iconPath,0"
$Shortcut.Description = "Sysgen RPO Loans Intelligence Suite"
$Shortcut.Save()

Write-Host "Created direct EXE Desktop Shortcut at: $shortcutPath"
