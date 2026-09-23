# Outlook 本机邮件监控器

这个版本只在本机读取经典版 Outlook 的收件箱，不连接 ChatGPT、不上传邮件、不需要 Microsoft 365 管理员授权。

## 使用方法

1. 第一次使用时复制 `config.json` 为 `config.local.json`，之后编辑本机的 `config.local.json`。
2. 在 `keywords` 中填写要监控的人名、别名或项目关键词；也可以在 `emailAddresses` 中填写精确邮箱地址。`config.local.json` 和运行状态文件不会被 Git 提交。
3. 根据需要修改 `intervalMinutes`，默认每 15 分钟检查一次。
4. 先手动测试：

   ```powershell
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\Monitor-OutlookMail.ps1
   ```

5. 测试成功后安装 Windows 定时任务：

   ```powershell
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\Install-OutlookMailMonitor.ps1
   ```

6. 停止监控：

   ```powershell
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\Uninstall-OutlookMailMonitor.ps1
   ```

## 说明

- 只读取收件箱，不标记已读、不移动、不删除、不发送邮件。
- 用 Outlook 邮件 EntryID 去重，重复运行不会反复提醒同一封邮件。
- 监控依赖经典版 Outlook 和 Windows 登录会话；新版 Outlook 可能不支持这种本地 COM 读取方式。
- 首次运行默认检查最近 60 分钟的邮件，状态保存在同目录的 `state.json`。
- 如果需要重新扫描最近窗口，可以删除 `state.json`，或运行：

  ```powershell
  pwsh -NoProfile -ExecutionPolicy Bypass -File .\Monitor-OutlookMail.ps1 -ResetState
  ```

- 通知默认只显示发件人、主题和命中原因，不显示完整正文。
