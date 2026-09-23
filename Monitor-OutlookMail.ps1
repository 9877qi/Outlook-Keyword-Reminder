[CmdletBinding()]
param(
    [string]$ConfigPath = $(if (Test-Path (Join-Path $PSScriptRoot 'config.local.json')) { Join-Path $PSScriptRoot 'config.local.json' } else { Join-Path $PSScriptRoot 'config.json' }),
    [switch]$ResetState,
    [switch]$TestNotification
)

$ErrorActionPreference = 'Stop'
$statePath = Join-Path $PSScriptRoot 'state.json'

function Read-JsonFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return $raw | ConvertFrom-Json
}

function Save-JsonFile([string]$Path, $Value) {
    $Value | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Get-SmtpSenderAddress($MailItem) {
    try {
        if ($MailItem.SenderEmailType -eq 'EX') {
            $exchangeUser = $MailItem.Sender.GetExchangeUser()
            if ($exchangeUser -and $exchangeUser.PrimarySmtpAddress) {
                return [string]$exchangeUser.PrimarySmtpAddress
            }
        }
    } catch { }
    try { return [string]$MailItem.SenderEmailAddress } catch { return '' }
}

function Convert-ToSimplifiedChinese([string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    if (-not ('OutlookKeywordReminder.ChineseText' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace OutlookKeywordReminder {
    public static class ChineseText {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int LCMapStringEx(
            string localeName, uint mapFlags, string source, int sourceCount,
            StringBuilder destination, int destinationCount,
            IntPtr versionInformation, IntPtr reserved, IntPtr sortHandle);

        public static string ToSimplified(string source) {
            const uint LCMAP_SIMPLIFIED_CHINESE = 0x02000000;
            int required = LCMapStringEx("zh-CN", LCMAP_SIMPLIFIED_CHINESE,
                source, source.Length, null, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            if (required == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
            var result = new StringBuilder(required);
            int written = LCMapStringEx("zh-CN", LCMAP_SIMPLIFIED_CHINESE,
                source, source.Length, result, result.Capacity,
                IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            if (written == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
            return result.ToString(0, written);
        }
    }
}
'@ -ErrorAction Stop
    }
    return [OutlookKeywordReminder.ChineseText]::ToSimplified($Text)
}

function Show-Notification([string]$Title, [string]$Message) {
    try {
        $toastType = [type]::GetType('Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime')
        $xmlType = [type]::GetType('Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType=WindowsRuntime')
        if ($toastType -and $xmlType) {
            $xml = [Activator]::CreateInstance($xmlType)
            $safeTitle = [System.Security.SecurityElement]::Escape($Title)
            $safeMessage = [System.Security.SecurityElement]::Escape($Message)
            $xml.LoadXml("<toast><visual><binding template='ToastGeneric'><text>$safeTitle</text><text>$safeMessage</text></binding></visual></toast>")
            $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
            $notifier = $toastType::CreateToastNotifier('Outlook Mail Monitor')
            $notifier.Show($toast)
            return
        }
    } catch { }

    try {
        $userName = [Environment]::UserName
        Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\msg.exe') -ArgumentList @($userName, "$Title`n$Message") -WindowStyle Hidden
    } catch {
        Write-Warning "无法显示 Windows 通知：$($_.Exception.Message)"
    }
}

if ($TestNotification) {
    Show-Notification 'Outlook 邮件监控测试' '这是一条测试提醒。监控规则命中时，也会以桌面通知提示。'
    Write-Output '已发送测试通知。'
    exit 0
}

$config = Read-JsonFile $ConfigPath
if (-not $config) { throw "找不到配置文件：$ConfigPath" }

$keywords = @($config.keywords | Where-Object { $_ -and $_.ToString().Trim() -and $_ -notlike '请在这里填写*' } | ForEach-Object { $_.ToString().Trim() })
$addresses = @($config.emailAddresses | Where-Object { $_ -and $_.ToString().Trim() } | ForEach-Object { $_.ToString().Trim().ToLowerInvariant() })
if ($keywords.Count -eq 0 -and $addresses.Count -eq 0) {
    throw '请先在 config.json 中填写 keywords 或 emailAddresses。'
}

$state = if ($ResetState) { $null } else { Read-JsonFile $statePath }
$now = Get-Date
$initialLookback = [Math]::Max(1, [int]$config.initialLookbackMinutes)
$overlap = [Math]::Max(0, [int]$config.overlapMinutes)
$lastChecked = if ($state -and $state.lastChecked) { [datetime]$state.lastChecked } else { $now.AddMinutes(-$initialLookback) }
$scanFrom = $lastChecked.AddMinutes(-$overlap)
$seen = [System.Collections.Generic.HashSet[string]]::new()
if ($state -and $state.seenEntryIds) { @($state.seenEntryIds) | ForEach-Object { [void]$seen.Add([string]$_) } }

$outlook = $null
$namespace = $null
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $folder = $namespace.GetDefaultFolder(6)
    $items = $folder.Items
    $items.Sort('[ReceivedTime]', $true)
    $inspected = 0
    $matched = 0
    $maxItems = [Math]::Max(1, [int]$config.maxItemsToInspect)

    foreach ($mail in $items) {
        if ($inspected -ge $maxItems) { break }
        $inspected++
        try {
            $received = [datetime]$mail.ReceivedTime
            if ($received -lt $scanFrom) { break }
            if ($received -gt $now) { continue }
            $entryId = [string]$mail.EntryID
            if ([string]::IsNullOrWhiteSpace($entryId) -or $seen.Contains($entryId)) { continue }

            $sender = Get-SmtpSenderAddress $mail
            $senderName = [string]$mail.SenderName
            $subject = [string]$mail.Subject
            $recipients = "{0} {1}" -f ([string]$mail.To), ([string]$mail.CC)
            $body = if ($config.matchBody) { [string]$mail.Body } else { '' }
            $senderForMatch = (Convert-ToSimplifiedChinese ($senderName + ' ' + $sender)).ToLowerInvariant()
            $subjectForMatch = (Convert-ToSimplifiedChinese $subject).ToLowerInvariant()
            $recipientsForMatch = (Convert-ToSimplifiedChinese $recipients).ToLowerInvariant()
            $bodyForMatch = (Convert-ToSimplifiedChinese $body).ToLowerInvariant()

            $matchReasons = [System.Collections.Generic.List[string]]::new()
            foreach ($address in $addresses) {
                if ($sender.ToLowerInvariant().Contains($address)) { [void]$matchReasons.Add("发件人：$address") }
                elseif ($recipients.ToLowerInvariant().Contains($address)) { [void]$matchReasons.Add("收件人：$address") }
            }
            foreach ($keyword in $keywords) {
                $needle = (Convert-ToSimplifiedChinese $keyword).ToLowerInvariant()
                if ($config.matchSender -and $senderForMatch.Contains($needle)) { [void]$matchReasons.Add("发件人提到：$keyword") }
                elseif ($config.matchSubject -and $subjectForMatch.Contains($needle)) { [void]$matchReasons.Add("主题包含：$keyword") }
                elseif ($config.matchRecipients -and $recipientsForMatch.Contains($needle)) { [void]$matchReasons.Add("收件人包含：$keyword") }
                elseif ($config.matchBody -and $bodyForMatch.Contains($needle)) { [void]$matchReasons.Add("正文包含：$keyword") }
            }

            [void]$seen.Add($entryId)
            if ($matchReasons.Count -gt 0) {
                $matched++
                $fromDisplay = if ($senderName) { "$senderName <$sender>" } else { $sender }
                $reason = ($matchReasons | Select-Object -Unique) -join '；'
                Show-Notification 'Outlook 邮件命中监控规则' "$fromDisplay`n$subject`n$reason"
                Write-Output ("命中：{0} | {1} | {2}" -f $fromDisplay, $subject, $reason)
            }
        } catch {
            Write-Warning "跳过一封无法读取的项目：$($_.Exception.Message)"
        }
    }

    $retained = @($seen | Select-Object -Last 1000)
    Save-JsonFile $statePath ([ordered]@{ lastChecked = $now.ToString('o'); seenEntryIds = $retained })
    Write-Output ("检查完成：检查 {0} 封，命中 {1} 封，时间 {2}" -f $inspected, $matched, $now.ToString('yyyy-MM-dd HH:mm:ss'))
}
finally {
    if ($namespace) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) }
    if ($outlook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
