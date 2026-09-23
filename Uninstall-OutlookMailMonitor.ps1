[CmdletBinding()]
param()
$taskName = 'Outlook Mail Monitor - Local'
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
Write-Output "已移除定时任务：$taskName"
