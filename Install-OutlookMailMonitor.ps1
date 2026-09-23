[CmdletBinding()]
param(
    [string]$MonitorDirectory = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$localConfigPath = Join-Path $MonitorDirectory 'config.local.json'
$configPath = if (Test-Path -LiteralPath $localConfigPath) { $localConfigPath } else { Join-Path $MonitorDirectory 'config.json' }
$scriptPath = Join-Path $MonitorDirectory 'Monitor-OutlookMail.ps1'
if (-not (Test-Path -LiteralPath $configPath)) { throw "找不到配置文件：$configPath" }
if (-not (Test-Path -LiteralPath $scriptPath)) { throw "找不到监控脚本：$scriptPath" }

$config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
$minutes = [Math]::Max(5, [int]$config.intervalMinutes)
$taskName = 'Outlook Mail Monitor - Local'
$powerShellPath = Join-Path $PSHOME 'pwsh.exe'
$action = New-ScheduledTaskAction -Execute $powerShellPath -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ConfigPath `"$configPath`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes $minutes) -RepetitionDuration (New-TimeSpan -Days 3650)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 4)
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Write-Output "已安装定时任务：$taskName；间隔：$minutes 分钟。"
Write-Output '前提：经典版 Outlook 已配置账号，并且用户已登录 Windows。'
