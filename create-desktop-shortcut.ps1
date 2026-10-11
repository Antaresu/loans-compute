$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "Sysgen RPO Loans.lnk"
$targetPath = "c:\Users\teyo\Downloads\Sysgen RPO loans compute\Sysgen-Loans-App.bat"
$workingDir = "c:\Users\teyo\Downloads\Sysgen RPO loans compute"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $targetPath
$Shortcut.WorkingDirectory = $workingDir
$Shortcut.Description = "Sysgen RPO Loans Intelligence Suite"
$Shortcut.Save()

Write-Host "Created shortcut successfully at: $shortcutPath"
