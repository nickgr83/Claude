$AppDir    = Split-Path -Parent $MyInvocation.MyCommand.Definition
$AppScript = Join-Path $AppDir "app.py"
$TaskName  = "AzureMessageCenterMonitor"

$PythonWCmd = Get-Command pythonw.exe -ErrorAction SilentlyContinue
if ($PythonWCmd) {
    $PythonW = $PythonWCmd.Source
} else {
    $PythonCmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if (-not $PythonCmd) {
        Write-Error "Python not found on PATH. Install Python and try again."
        exit 1
    }
    $PythonW = $PythonCmd.Source
    Write-Warning "pythonw.exe not found; using python.exe (console window may flash at login)."
}

Write-Host "Using: $PythonW"
Write-Host "Script: $AppScript"

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$Action   = New-ScheduledTaskAction -Execute $PythonW -Argument "`"$AppScript`"" -WorkingDirectory $AppDir
$Trigger  = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action   $Action `
    -Trigger  $Trigger `
    -Settings $Settings `
    -RunLevel Limited `
    -Force | Out-Null

Write-Host ""
Write-Host "Task registered. The app will start automatically at next login."
Write-Host ""
Write-Host "To start it now without rebooting:"
Write-Host "  Start-ScheduledTask -TaskName AzureMessageCenterMonitor"
Write-Host ""
Write-Host "To remove: run uninstall_startup.ps1"
