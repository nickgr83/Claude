$TaskName = "AzureMessageCenterMonitor"

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if (-not $task) {
    Write-Host "Task '$TaskName' not found — nothing to remove."
    exit 0
}

Stop-ScheduledTask  -TaskName $TaskName -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false

Write-Host "✓ Task '$TaskName' removed. The app will no longer start at login."
