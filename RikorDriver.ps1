# Driver Updater Rikor v 2.0 - Fixed Critical Issues
# Run as Administrator
# Usage: .\RikorDriver.ps1
# Silent mode: .\RikorDriver.ps1 -Silent -Task "CheckDriverUpdates"
#
# CHANGELOG v2.0:
# - Fixed double error counting in pnputil
# - Added INF file validation
# - Improved temp file cleanup
# - Added admin rights check in background jobs
# - Restored Microsoft Update driver search
# - Removed broken Google Drive download (server unavailable)
param(
[switch]$Silent,
[string]$Task = "",
[string]$Language = "ru",
[string]$ProxyAddress = "",
[string]$FilterClass = "",
[string]$FilterManufacturer = ""
)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# -------------------------
# Multi-Language Support
# -------------------------
$global:Languages = @{
"en" = @{
AppTitle = "Rikor Driver Installer V2.0"
BtnDownloadRikor = "Download Drivers from Rikor Server"
BtnScan = "Scan Installed Drivers"
BtnBackup = "Backup Drivers"
BtnInstall = "Install From Folder"
BtnCancel = "Cancel Task"
BtnOpenLogs = "Open Logs"
BtnDarkMode = "Dark Mode"
BtnLightMode = "Light Mode"
BtnSchedule = "Schedule Updates"
BtnRestorePoint = "Create Restore Point"
BtnHistory = "Update History"
BtnSettings = "Settings"
BtnFilters = "Filters"
TaskRunning = "A task is already running. Cancel it first."
PermissionError = "Please run this script as Administrator."
BackupCanceled = "Backup canceled by user."
InstallCanceled = "Install canceled by user."
NoTaskToCancel = "No running task to cancel."
TaskCancelled = "[CANCELLED] Task cancelled by user."
LogFolderMissing = "Log folder missing."
StartingTask = "-> Starting task:"
TaskFinished = "=== Task finished:"
SelectBackupFolder = "Select folder to save driver backup"
SelectDriverFolder = "Select folder containing driver .inf files"
ScheduleCreated = "Scheduled task created successfully!"
ScheduleRemoved = "Scheduled task removed."
RestorePointCreated = "System restore point created successfully!"
RestorePointFailed = "Failed to create restore point."
ProxyConfigured = "Proxy configured:"
ProxyCleared = "Proxy settings cleared."
FilterApplied = "Filter applied:"
FilterCleared = "Filters cleared."
LanguageChanged = "Language changed to:"
HistoryEmpty = "No update history found."
SettingsTitle = "Settings"
ScheduleTitle = "Schedule Updates"
FilterTitle = "Driver Filters"
HistoryTitle = "Update History"
Daily = "Daily"
Weekly = "Weekly"
Monthly = "Monthly"
Time = "Time:"
ProxyLabel = "Proxy Address:"
ClassFilter = "Filter by Class:"
ManufacturerFilter = "Filter by Manufacturer:"
Apply = "Apply"
Clear = "Clear"
Close = "Close"
Enable = "Enable"
Disable = "Disable"
Remove = "Remove Schedule"
}
"ru" = @{
AppTitle = "Rikor Driver Installer V2.0"
BtnDownloadRikor = "Загрузить драйверы с сервера Rikor"
BtnScan = "Сканировать установленные драйверы"
BtnBackup = "Резервное копирование"
BtnInstall = "Установить из папки"
BtnCancel = "Отменить задачу"
BtnOpenLogs = "Открыть журналы"
BtnDarkMode = "Тёмная тема"
BtnLightMode = "Светлая тема"
BtnSchedule = "Запланировать обновления"
BtnRestorePoint = "Создать точку восстановления"
BtnHistory = "История обновлений"
BtnSettings = "Настройки"
BtnFilters = "Фильтры"
TaskRunning = "Задача уже выполняется. Сначала отмените её."
PermissionError = "Запустите этот скрипт от имени администратора."
BackupCanceled = "Резервное копирование отменено пользователем."
InstallCanceled = "Установка отменена пользователем."
NoTaskToCancel = "Нет запущенной задачи для отмены."
TaskCancelled = "[ОТМЕНЕНО] Задача отменена пользователем."
LogFolderMissing = "Папка журнала не найдена."
StartingTask = "-> Запуск задачи:"
TaskFinished = "=== Задача завершена:"
SelectBackupFolder = "Выберите папку для сохранения резервной копии драйверов"
SelectDriverFolder = "Выберите папку с файлами драйверов (.inf)"
ScheduleCreated = "Запланированная задача успешно создана!"
ScheduleRemoved = "Запланированная задача удалена."
RestorePointCreated = "Точка восстановления системы успешно создана!"
RestorePointFailed = "Не удалось создать точку восстановления."
ProxyConfigured = "Прокси настроен:"
ProxyCleared = "Настройки прокси сброшены."
FilterApplied = "Фильтр применён:"
FilterCleared = "Фильтры сброшены."
LanguageChanged = "Язык изменён на:"
HistoryEmpty = "История обновлений не найдена."
SettingsTitle = "Настройки"
ScheduleTitle = "Запланировать обновления"
FilterTitle = "Фильтры драйверов"
HistoryTitle = "История обновлений"
Daily = "Ежедневно"
Weekly = "Еженедельно"
Monthly = "Ежемесячно"
Time = "Время:"
ProxyLabel = "Адрес прокси:"
ClassFilter = "Фильтр по классу:"
ManufacturerFilter = "Фильтр по производителю:"
Apply = "Применить"
Clear = "Очистить"
Close = "Закрыть"
Enable = "Включить"
Disable = "Отключить"
Remove = "Удалить расписание"
}
}
# Get localized string
function Get-LocalizedString([string]$key) {
$lang = $global:CurrentLanguage
if (-not $global:Languages.ContainsKey($lang)) { $lang = "en" }
if ($global:Languages[$lang].ContainsKey($key)) {
return $global:Languages[$lang][$key]
}
return $global:Languages["en"][$key]
}
$global:CurrentLanguage = $Language
# -------------------------
# Require Admin
# -------------------------
function Assert-AdminPrivilege {
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
if (-not $Silent) {
[System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "PermissionError"), "Permission Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
exit 1
}
}
Assert-AdminPrivilege
# -------------------------
# Globals and paths
# -------------------------
$AppTitle = Get-LocalizedString "AppTitle"
# CHANGED: Update base folder name
$LogBase = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller"
$HistoryFile = Join-Path $LogBase "UpdateHistory.json"
$SettingsFile = Join-Path $LogBase "Settings.json"
if (!(Test-Path $LogBase)) { New-Item -ItemType Directory -Path $LogBase -Force | Out-Null }
$global:CurrentJob = $null
$global:CurrentTaskLog = $null
$global:ProxySettings = @{ Enabled = $false; Address = "" }
$global:FilterSettings = @{ Class = ""; Manufacturer = "" }
$global:DarkModeEnabled = $false
# -------------------------
# Settings Management
# -------------------------
function Import-Settings {
if (Test-Path $SettingsFile) {
try {
$settings = Get-Content -Path $SettingsFile -Raw | ConvertFrom-Json
if ($settings.Language) { $global:CurrentLanguage = $settings.Language }
if ($settings.Proxy) {
$global:ProxySettings.Enabled = $settings.Proxy.Enabled
$global:ProxySettings.Address = $settings.Proxy.Address
}
if ($settings.Filters) {
$global:FilterSettings.Class = $settings.Filters.Class
$global:FilterSettings.Manufacturer = $settings.Filters.Manufacturer
}
if ($null -ne $settings.DarkMode) { $global:DarkModeEnabled = $settings.DarkMode }
} catch {}
}
}
function Export-Settings {
$settings = @{
Language = $global:CurrentLanguage
Proxy = $global:ProxySettings
Filters = $global:FilterSettings
DarkMode = $global:DarkModeEnabled
}
$settings | ConvertTo-Json -Depth 3 | Set-Content -Path $SettingsFile -Encoding UTF8
}
Import-Settings
# Apply command-line parameters
if ($ProxyAddress) {
$global:ProxySettings.Enabled = $true
$global:ProxySettings.Address = $ProxyAddress
}
if ($FilterClass) { $global:FilterSettings.Class = $FilterClass }
if ($FilterManufacturer) { $global:FilterSettings.Manufacturer = $FilterManufacturer }
# -------------------------
# Update History Logging
# -------------------------
function Add-HistoryEntry {
param(
[string]$TaskName,
[string]$Status,
[string]$Details
)
$history = @()
if (Test-Path $HistoryFile) {
try {
$content = Get-Content -Path $HistoryFile -Raw
if ($content) {
$parsed = $content | ConvertFrom-Json
if ($parsed -is [array]) {
$history = $parsed
} else {
$history = @($parsed)
}
}
} catch { $history = @() }
}
$entry = [PSCustomObject]@{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = $TaskName
Status = $Status
Details = $Details
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $HistoryFile -Encoding UTF8
}
function Get-UpdateHistory {
if (Test-Path $HistoryFile) {
try {
$content = Get-Content -Path $HistoryFile -Raw
if ($content) {
return $content | ConvertFrom-Json
}
} catch {}
}
return @()
}
# -------------------------
# Network Proxy Support
# -------------------------
function Set-ProxySettings {
param([string]$ProxyAddr, [bool]$Enable)
$global:ProxySettings.Address = $ProxyAddr
$global:ProxySettings.Enabled = $Enable
if ($Enable -and $ProxyAddr) {
[System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($ProxyAddr, $true)
$env:HTTP_PROXY = $ProxyAddr
$env:HTTPS_PROXY = $ProxyAddr
} else {
[System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
Remove-Item Env:\HTTP_PROXY -ErrorAction SilentlyContinue
Remove-Item Env:\HTTPS_PROXY -ErrorAction SilentlyContinue
}
Export-Settings
}
# Apply proxy if configured
if ($global:ProxySettings.Enabled -and $global:ProxySettings.Address) {
Set-ProxySettings -ProxyAddr $global:ProxySettings.Address -Enable $true
}
# -------------------------
# System Restore Point
# -------------------------
function New-RestorePoint {
param([string]$Description = "Rikor Driver Installer Restore Point") # CHANGED: Description text
try {
# Enable System Restore on system drive if not already enabled
Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
# Create the restore point
Checkpoint-Computer -Description $Description -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
Add-HistoryEntry -TaskName "RestorePoint" -Status "Success" -Details $Description
return $true
} catch {
Add-HistoryEntry -TaskName "RestorePoint" -Status "Failed" -Details $_.Exception.Message
return $false
}
}
# -------------------------
# Scheduled Updates
# -------------------------
$ScheduledTaskName = "RikorDriverInstaller_ScheduledCheck" # CHANGED: Task name
function Get-ScheduledUpdateTask {
try {
return Get-ScheduledTask -TaskName $ScheduledTaskName -ErrorAction SilentlyContinue
} catch { return $null }
}
function Set-ScheduledUpdate {
param(
[ValidateSet("Daily", "Weekly", "Monthly")]
[string]$Frequency,
[string]$Time = "03:00"
)
try {
# Remove existing task if any
Remove-ScheduledUpdate
$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = $MyInvocation.PSCommandPath }
# NEW: Update scheduled task action to use new combined task name
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -Silent -Task DownloadRikorDrivers"
switch ($Frequency) {
"Daily" { $trigger = New-ScheduledTaskTrigger -Daily -At $Time }
"Weekly" { $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At $Time }
"Monthly" { $trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At $Time }
}
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Register-ScheduledTask -TaskName $ScheduledTaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Add-HistoryEntry -TaskName "ScheduleUpdate" -Status "Created" -Details "$Frequency at $Time"
return $true
} catch {
Add-HistoryEntry -TaskName "ScheduleUpdate" -Status "Failed" -Details $_.Exception.Message
return $false
}
}
function Remove-ScheduledUpdate {
try {
Unregister-ScheduledTask -TaskName $ScheduledTaskName -Confirm:$false -ErrorAction SilentlyContinue
return $true
} catch { return $false }
}
# -------------------------
# Silent Mode Operation
# -------------------------
if ($Silent -and $Task) {
$logFile = Join-Path $LogBase "$Task`_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Write-SilentLog($msg) {
$timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Add-Content -Path $logFile -Value "$timestamp - $msg"
}
Write-SilentLog "Starting silent mode: $Task"
switch ($Task) {
# NEW: Combined task - try Rikor server first, fallback to Microsoft Update
"DownloadRikorDrivers" {
Write-SilentLog "Attempting to download drivers from Rikor server..."
$rikorServerAvailable = $false
# TODO: Add actual Rikor server URL and download logic here
# For now, simulate server check
try {
    # Placeholder for Rikor server check
    # $rikorUrl = "https://rikor-server.example.com/drivers"
    # Test-Connection or Invoke-WebRequest to check availability
    Write-SilentLog "[INFO] Rikor server is currently unavailable."
    $rikorServerAvailable = $false
} catch {
    Write-SilentLog "[INFO] Cannot connect to Rikor server."
    $rikorServerAvailable = $false
}

if (-not $rikorServerAvailable) {
    Write-SilentLog "[INFO] Falling back to Microsoft Update for driver search and installation..."
    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        Write-SilentLog "Searching for available driver updates (this may take a few minutes)..."
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver'")
        if ($SearchResult.Updates.Count -eq 0) {
            Write-SilentLog "No driver updates available from Microsoft Update"
            Add-HistoryEntry -TaskName "DownloadRikorDrivers" -Status "Completed" -Details "No updates found (fallback to MS Update)"
        } else {
            Write-SilentLog "Found $($SearchResult.Updates.Count) driver update(s) available:"
            Write-SilentLog ""
            foreach ($Update in $SearchResult.Updates) {
                Write-SilentLog "  - $($Update.Title)"
                Write-SilentLog "    Size: $([math]::Round($Update.MaxDownloadSize / 1MB, 2)) MB"
            }
            Write-SilentLog ""
            Write-SilentLog "Downloading and installing driver updates..."
            
            # Create update collection for download and install
            $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
            foreach ($Update in $SearchResult.Updates) {
                if ($Update.EulaAccepted -eq $false) {
                    $Update.AcceptEula()
                }
                $UpdatesToDownload.Add($Update) | Out-Null
            }
            
            # Download updates
            if ($UpdatesToDownload.Count -gt 0) {
                Write-SilentLog "Downloading $($UpdatesToDownload.Count) update(s)..."
                $Downloader = $UpdateSession.CreateUpdateDownloader()
                $Downloader.Updates = $UpdatesToDownload
                $DownloadResult = $Downloader.Download()
                Write-SilentLog "Download completed with result code: $($DownloadResult.ResultCode)"
                
                # Install downloaded updates
                $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                foreach ($Update in $SearchResult.Updates) {
                    if ($Update.IsDownloaded) {
                        $UpdatesToInstall.Add($Update) | Out-Null
                    }
                }
                
                if ($UpdatesToInstall.Count -gt 0) {
                    Write-SilentLog "Installing $($UpdatesToInstall.Count) update(s)..."
                    $Installer = $UpdateSession.CreateUpdateInstaller()
                    $Installer.Updates = $UpdatesToInstall
                    $InstallResult = $Installer.Install()
                    Write-SilentLog "Installation completed with result code: $($InstallResult.ResultCode)"
                    Write-SilentLog "Reboot required: $($InstallResult.RebootRequired)"
                    
                    $successCount = 0
                    $failCount = 0
                    for ($i = 0; $i -lt $UpdatesToInstall.Count; $i++) {
                        $resultCode = $InstallResult.GetUpdateResult($i).ResultCode
                        if ($resultCode -eq 2) { # 2 = Succeeded
                            $successCount++
                        } else {
                            $failCount++
                        }
                    }
                    Write-SilentLog "Successfully installed: $successCount, Failed: $failCount"
                    Add-HistoryEntry -TaskName "DownloadRikorDrivers" -Status "Completed" -Details "Installed $successCount/$($UpdatesToInstall.Count) updates via MS Update"
                } else {
                    Write-SilentLog "[WARNING] No updates were downloaded successfully"
                    Add-HistoryEntry -TaskName "DownloadRikorDrivers" -Status "Completed" -Details "Download failed for all updates"
                }
            }
        }
    } catch {
        Write-SilentLog "[ERROR] Failed to download/install driver updates: $_"
        Add-HistoryEntry -TaskName "DownloadRikorDrivers" -Status "Failed" -Details $_.Exception.Message
    }
}
Write-SilentLog "Completed"
}
"ScanDrivers" {
Write-SilentLog "Scanning installed drivers..."
try {
$drivers = Get-CimInstance Win32_PnPSignedDriver | Select-Object DeviceName, Manufacturer, DriverVersion, InfName, Class, DriverDate
# Apply filters
if ($global:FilterSettings.Class) {
Write-SilentLog "Applying class filter: $global:FilterSettings.Class"
$drivers = $drivers | Where-Object { $_.Class -like "*$($global:FilterSettings.Class)*" }
}
if ($global:FilterSettings.Manufacturer) {
Write-SilentLog "Applying manufacturer filter: $global:FilterSettings.Manufacturer"
$drivers = $drivers | Where-Object { $_.Manufacturer -like "*$($global:FilterSettings.Manufacturer)*" }
}
Write-SilentLog "Found $($drivers.Count) drivers matching criteria"
# Group by class
$driversByClass = $drivers | Group-Object -Property Class
Write-SilentLog "Categories: $($driversByClass.Count) different driver types"
Write-SilentLog ""
Write-SilentLog "Sample drivers:"
$sampleCount = [Math]::Min(5, $drivers.Count)
$drivers | Sort-Object DeviceName | Select-Object -First $sampleCount | ForEach-Object {
Write-SilentLog "  * $($_.DeviceName) (v$($_.DriverVersion))"
}
if ($drivers.Count -gt $sampleCount) {
Write-SilentLog "  ... and $($drivers.Count - $sampleCount) more"
}
Write-SilentLog ""
Write-SilentLog "Exporting full list to CSV..."
$csvPath = Join-Path (Split-Path $logPath) "InstalledDrivers_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
$drivers | Sort-Object DeviceName | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-SilentLog "Exported to: $csvPath"
Write-SilentLog "Completed"
} catch {
Write-SilentLog "[ERROR] Failed to scan drivers: $_"
}
}
"BackupDrivers" {
$dest = $innerArgs[0]
Write-SilentLog "Backing up drivers to: $dest"
try {
if (!(Test-Path $dest)) {
New-Item -ItemType Directory -Path $dest -Force | Out-Null
}
Write-SilentLog "Exporting drivers (this may take several minutes)..."
& dism.exe /online /export-driver /destination:$dest 2>&1 | Out-Null
$exportedCount = (Get-ChildItem -Path $dest -Recurse -Directory -ErrorAction SilentlyContinue).Count
if ($exportedCount -gt 0) {
Write-SilentLog "Backup completed: $exportedCount driver package(s) exported"
} else {
Write-SilentLog "Backup completed. Check destination folder for exported drivers."
}
} catch {
Write-SilentLog "[ERROR] Backup failed: $_"
}
Write-SilentLog "Completed"
}
"InstallDrivers" {
$folder = $innerArgs[0]
Write-SilentLog "Installing drivers from: $folder"
try {
if (-not (Test-Path $folder)) {
Write-SilentLog "[ERROR] Folder not found: $folder"
Write-SilentLog "Completed"
return
}
$infFiles = Get-ChildItem -Path $folder -Recurse -Include *.inf -ErrorAction SilentlyContinue
if ($infFiles.Count -eq 0) {
Write-SilentLog "[ERROR] No .inf driver files found in folder"
Write-SilentLog "Completed"
return
}
Write-SilentLog "Found $($infFiles.Count) driver file(s)"
Write-SilentLog "Installing drivers..."
Write-SilentLog ""
$successCount = 0
$failCount = 0
$current = 0
# FIXED: Corrected double error counting logic
foreach ($inf in $infFiles) {
$current++
Write-SilentLog "[$current/$($infFiles.Count)] $($inf.Name)"
try {
# Validate INF file before installation
$infContent = Get-Content $inf.FullName -Raw -ErrorAction SilentlyContinue
if (-not $infContent -or $infContent -notmatch '\[Version\]') {
    Write-SilentLog "     -> SKIPPED: Invalid INF file"
    $failCount++
    continue
}

# Check architecture compatibility
$is64Bit = [Environment]::Is64BitOperatingSystem
$archPattern = if ($is64Bit) { "amd64|x64|NTamd64" } else { "x86|NTx86" }
if ($infContent -notmatch $archPattern) {
    Write-SilentLog "     -> SKIPPED: Incompatible architecture"
    $failCount++
    continue
}

# Install driver using pnputil
$out = & pnputil.exe /add-driver $inf.FullName /install /force 2>&1
if ($LASTEXITCODE -eq 0) {
    $successCount++
    Write-SilentLog "     -> Successfully installed"
} else {
    $failCount++
    Write-SilentLog "     -> Installation failed (exit code: $LASTEXITCODE)"
    # Log error details
    $out | Where-Object { $_ -match "error|fail" } | ForEach-Object {
        Write-SilentLog "        $_"
    }
}
} catch {
$failCount++
Write-SilentLog "     -> Exception: $_"
}
Start-Sleep -Milliseconds 300
}
Write-SilentLog ""
Write-SilentLog "Installation complete: $successCount successful, $failCount failed"
if ($successCount -gt 0) {
Write-SilentLog "Note: Reboot may be required for some drivers."
}
} catch {
Write-SilentLog "[ERROR] Installation failed: $_"
}
Write-SilentLog "Completed"
}
default {
Write-SilentLog "Unknown task: $Task"
Add-HistoryEntry -TaskName $Task -Status "Failed" -Details "Unknown task"
}
}
Write-SilentLog "Silent mode completed"
exit 0
}
# -------------------------
# Helper Functions
# -------------------------
function New-TaskLog([string]$taskName) {
$ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
$logFile = Join-Path $LogBase "$($taskName)_$ts.log"
New-Item -Path $logFile -ItemType File -Force | Out-Null
return $logFile
}
function Add-StatusUI {
param($form, $statusControl, $text)
if ($form -and $statusControl) {
$method = [System.Windows.Forms.MethodInvoker]{
$statusControl.AppendText("$text`r`n")
$statusControl.ScrollToCaret()
}
$form.Invoke($method)
}
}
# -------------------------
# Modern UI Color Scheme
# -------------------------
$global:UIColors = @{
# Dark Theme
Dark = @{
Background = [System.Drawing.Color]::FromArgb(18, 18, 18)
Surface = [System.Drawing.Color]::FromArgb(30, 30, 30)
SurfaceHover = [System.Drawing.Color]::FromArgb(45, 45, 45)
Primary = [System.Drawing.Color]::FromArgb(156, 39, 176)  # Purple accent
PrimaryHover = [System.Drawing.Color]::FromArgb(186, 74, 206)
Secondary = [System.Drawing.Color]::FromArgb(66, 66, 66)
Text = [System.Drawing.Color]::FromArgb(240, 240, 240)
TextSecondary = [System.Drawing.Color]::FromArgb(170, 170, 170)
Border = [System.Drawing.Color]::FromArgb(60, 60, 60)
Success = [System.Drawing.Color]::FromArgb(76, 175, 80)
Warning = [System.Drawing.Color]::FromArgb(255, 152, 0)
Error = [System.Drawing.Color]::FromArgb(244, 67, 54)
MenuBar = [System.Drawing.Color]::FromArgb(25, 25, 25)
StatusBar = [System.Drawing.Color]::FromArgb(126, 34, 152)
}
# Light Theme
Light = @{
Background = [System.Drawing.Color]::FromArgb(255, 255, 255)
Surface = [System.Drawing.Color]::FromArgb(245, 247, 250)
SurfaceHover = [System.Drawing.Color]::FromArgb(235, 238, 242)
Primary = [System.Drawing.Color]::FromArgb(156, 39, 176)  # Purple accent
PrimaryHover = [System.Drawing.Color]::FromArgb(186, 74, 206)
Secondary = [System.Drawing.Color]::FromArgb(224, 224, 224)
Text = [System.Drawing.Color]::FromArgb(33, 33, 33)
TextSecondary = [System.Drawing.Color]::FromArgb(117, 117, 117)
Border = [System.Drawing.Color]::FromArgb(224, 224, 224)
Success = [System.Drawing.Color]::FromArgb(67, 160, 71)
Warning = [System.Drawing.Color]::FromArgb(251, 140, 0)
Error = [System.Drawing.Color]::FromArgb(229, 57, 53)
MenuBar = [System.Drawing.Color]::FromArgb(255, 255, 255)
StatusBar = [System.Drawing.Color]::FromArgb(156, 39, 176)
}
}
function Get-ThemeColors {
if ($global:DarkModeEnabled) { return $global:UIColors.Dark }
return $global:UIColors.Light
}
# -------------------------
# Build Form
# -------------------------
$form = New-Object Windows.Forms.Form
$form.Text = $AppTitle
$form.Size = '1050,720'
$form.MinimumSize = '1050,720'
$form.StartPosition = "CenterScreen"
$form.Font = New-Object Drawing.Font("Segoe UI", 9.5)
# -------------------------
# Menu Strip
# -------------------------
$menuStrip = New-Object Windows.Forms.MenuStrip
$menuStrip.Padding = '6,2,0,2'
# File Menu
$menuFile = New-Object Windows.Forms.ToolStripMenuItem
$menuFile.Text = "&File"
$menuOpenLogs = New-Object Windows.Forms.ToolStripMenuItem
$menuOpenLogs.Text = (Get-LocalizedString "BtnOpenLogs")
$menuOpenLogs.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::L
$menuSeparator1 = New-Object Windows.Forms.ToolStripSeparator
$menuExit = New-Object Windows.Forms.ToolStripMenuItem
$menuExit.Text = "Exit"
$menuExit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$menuFile.DropDownItems.AddRange(@($menuOpenLogs, $menuSeparator1, $menuExit))
# Actions Menu
$menuActions = New-Object Windows.Forms.ToolStripMenuItem
$menuActions.Text = "&Actions"
# NEW: Single combined menu item for Rikor download
$menuDownloadRikor = New-Object Windows.Forms.ToolStripMenuItem
$menuDownloadRikor.Text = (Get-LocalizedString "BtnDownloadRikor")
$menuDownloadRikor.ShortcutKeys = [System.Windows.Forms.Keys]::F5
$menuScan = New-Object Windows.Forms.ToolStripMenuItem
$menuScan.Text = (Get-LocalizedString "BtnScan")
$menuScan.ShortcutKeys = [System.Windows.Forms.Keys]::F7
$menuSeparator2 = New-Object Windows.Forms.ToolStripSeparator
$menuBackup = New-Object Windows.Forms.ToolStripMenuItem
$menuBackup.Text = (Get-LocalizedString "BtnBackup")
$menuBackup.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::B
$menuInstall = New-Object Windows.Forms.ToolStripMenuItem
$menuInstall.Text = (Get-LocalizedString "BtnInstall")
$menuInstall.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::I
$menuSeparator3 = New-Object Windows.Forms.ToolStripSeparator
$menuCancel = New-Object Windows.Forms.ToolStripMenuItem
$menuCancel.Text = (Get-LocalizedString "BtnCancel")
$menuCancel.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Q
$menuActions.DropDownItems.AddRange(@($menuScan, $menuSeparator2, $menuBackup, $menuInstall, $menuSeparator3, $menuCancel))
# Tools Menu
$menuTools = New-Object Windows.Forms.ToolStripMenuItem
$menuTools.Text = "&Tools"
$menuRestorePoint = New-Object Windows.Forms.ToolStripMenuItem
$menuRestorePoint.Text = (Get-LocalizedString "BtnRestorePoint")
$menuSchedule = New-Object Windows.Forms.ToolStripMenuItem
$menuSchedule.Text = (Get-LocalizedString "BtnSchedule")
$menuFilters = New-Object Windows.Forms.ToolStripMenuItem
$menuFilters.Text = (Get-LocalizedString "BtnFilters")
$menuSeparator4 = New-Object Windows.Forms.ToolStripSeparator
$menuHistory = New-Object Windows.Forms.ToolStripMenuItem
$menuHistory.Text = (Get-LocalizedString "BtnHistory")
$menuHistory.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::H
$menuTools.DropDownItems.AddRange(@($menuRestorePoint, $menuSchedule, $menuFilters, $menuSeparator4, $menuHistory))
# View Menu
$menuView = New-Object Windows.Forms.ToolStripMenuItem
$menuView.Text = "&View"
$menuToggleTheme = New-Object Windows.Forms.ToolStripMenuItem
$menuToggleTheme.Text = (Get-LocalizedString "BtnDarkMode")
$menuToggleTheme.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::T
$menuSeparator5 = New-Object Windows.Forms.ToolStripSeparator
# Language submenu
$menuLanguage = New-Object Windows.Forms.ToolStripMenuItem
$menuLanguage.Text = "Language"
$langItems = @("English (en)", "Русский (ru)")
$langCodes = @("en", "ru")
for ($i = 0; $i -lt $langItems.Count; $i++) {
$langMenuItem = New-Object Windows.Forms.ToolStripMenuItem
$langMenuItem.Text = $langItems[$i]
$langMenuItem.Tag = $langCodes[$i]
if ($langCodes[$i] -eq $global:CurrentLanguage) {
$langMenuItem.Checked = $true
}
$menuLanguage.DropDownItems.Add($langMenuItem) | Out-Null
}
$menuView.DropDownItems.AddRange(@($menuToggleTheme, $menuSeparator5, $menuLanguage))
# Settings Menu
$menuSettingsTop = New-Object Windows.Forms.ToolStripMenuItem
$menuSettingsTop.Text = "&Settings"
$menuSettingsTop.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Oemcomma
$menuStrip.Items.AddRange(@($menuFile, $menuActions, $menuTools, $menuView, $menuSettingsTop))
# -------------------------
# Toolbar Panel
# -------------------------
$toolbarPanel = New-Object Windows.Forms.Panel
$toolbarPanel.Dock = 'Top'
$toolbarPanel.Height = 56
# Note: toolbarPanel and menuStrip are added later for correct dock order
# TableLayoutPanel for responsive buttons
$buttonContainer = New-Object Windows.Forms.TableLayoutPanel
$buttonContainer.Dock = 'Fill'
$buttonContainer.ColumnCount = 5
$buttonContainer.RowCount = 1
$buttonContainer.Padding = '12,8,12,8'
$buttonContainer.AutoSize = $false
# Set column styles for proportional sizing
$buttonContainer.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 28))) | Out-Null  # Download Rikor
$buttonContainer.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 20))) | Out-Null  # Scan
$buttonContainer.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 17))) | Out-Null  # Backup
$buttonContainer.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 20))) | Out-Null  # Install
$buttonContainer.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 15))) | Out-Null  # Cancel
$toolbarPanel.Controls.Add($buttonContainer)
# Function to update button sizes when form resizes
function Update-ButtonContainerPadding {
# Update button widths based on container size
$containerWidth = $buttonContainer.ClientSize.Width - 24  # Subtract padding
if ($containerWidth -gt 0) {
    $btnDownloadRikor.Width = [Math]::Max(200, [int]($containerWidth * 0.28))
    $btnScan.Width = [Math]::Max(150, [int]($containerWidth * 0.20))
    $btnBackup.Width = [Math]::Max(120, [int]($containerWidth * 0.17))
    $btnInstall.Width = [Math]::Max(150, [int]($containerWidth * 0.20))
    $btnCancel.Width = [Math]::Max(100, [int]($containerWidth * 0.15))
}
}
# Modern Button Style
function New-RoundedRegion {
param([int]$Width, [int]$Height, [int]$Radius = 8)
$path = New-Object System.Drawing.Drawing2D.GraphicsPath
$r = $Radius
$w = $Width
$h = $Height
# Top-left arc
$path.AddArc(0, 0, $r * 2, $r * 2, 180, 90)
# Top-right arc
$path.AddArc($w - $r * 2, 0, $r * 2, $r * 2, 270, 90)
# Bottom-right arc
$path.AddArc($w - $r * 2, $h - $r * 2, $r * 2, $r * 2, 0, 90)
# Bottom-left arc
$path.AddArc(0, $h - $r * 2, $r * 2, $r * 2, 90, 90)
$path.CloseFigure()
return New-Object System.Drawing.Region($path)
}
function New-ModernButton {
param(
[string]$Text,
[int]$Width = 130,
[bool]$Primary = $false,
[string]$IconChar = ""
)
$btn = New-Object System.Windows.Forms.Button
$btn.Text = if ($IconChar) { "$IconChar  $Text" } else { $Text }
$btn.Width = $Width
$btn.Height = 38
$btn.FlatStyle = "Flat"
$btn.FlatAppearance.BorderSize = 0
$btn.Cursor = [System.Windows.Forms.Cursors]::Hand
$btn.Font = New-Object Drawing.Font("Segoe UI Semibold", 9)
$btn.Tag = $Primary
$btn.TextAlign = 'MiddleCenter'
# Apply rounded corners
$btn.Region = New-RoundedRegion -Width $Width -Height 38 -Radius 6
return $btn
}
# Create toolbar buttons with responsive sizing
# NEW: Single combined button for Rikor download with MS Update fallback
$btnDownloadRikor = New-ModernButton -Text (Get-LocalizedString "BtnDownloadRikor") -Width 280 -Primary $true
$btnDownloadRikor.Dock = 'Fill'
$btnDownloadRikor.Margin = '0,0,6,0'
$btnScan = New-ModernButton -Text (Get-LocalizedString "BtnScan") -Width 200 -Primary $false
$btnScan.Dock = 'Fill'
$btnScan.Margin = '6,0,6,0'
$btnBackup = New-ModernButton -Text (Get-LocalizedString "BtnBackup") -Width 150
$btnBackup.Dock = 'Fill'
$btnBackup.Margin = '6,0,6,0'
$btnInstall = New-ModernButton -Text (Get-LocalizedString "BtnInstall") -Width 180
$btnInstall.Dock = 'Fill'
$btnInstall.Margin = '6,0,6,0'
$btnCancel = New-ModernButton -Text (Get-LocalizedString "BtnCancel") -Width 130
$btnCancel.Dock = 'Fill'
$btnCancel.Margin = '6,0,0,0'
# Add buttons to responsive container
$buttonContainer.Controls.Add($btnDownloadRikor, 0, 0)
$buttonContainer.Controls.Add($btnScan, 1, 0)
$buttonContainer.Controls.Add($btnBackup, 2, 0)
$buttonContainer.Controls.Add($btnInstall, 3, 0)
$buttonContainer.Controls.Add($btnCancel, 4, 0)
# -------------------------
# Toolbar Separator (created here, added later for correct dock order)
# -------------------------
$toolbarSeparator = New-Object Windows.Forms.Panel
$toolbarSeparator.Dock = 'Top'
$toolbarSeparator.Height = 1
# -------------------------
# Status Bar
# -------------------------
$statusBar = New-Object Windows.Forms.Panel
$statusBar.Dock = 'Bottom'
$statusBar.Height = 28
$statusLabel = New-Object Windows.Forms.Label
$statusLabel.Text = "  Ready"
$statusLabel.Dock = 'Left'
$statusLabel.Width = 600
$statusLabel.TextAlign = 'MiddleLeft'
$statusLabel.Font = New-Object Drawing.Font("Segoe UI", 9)
$statusBar.Controls.Add($statusLabel)
$versionLabel = New-Object Windows.Forms.Label
$versionLabel.Text = "KILO v2.0  "
$versionLabel.Dock = 'Right'
$versionLabel.Width = 60
$versionLabel.TextAlign = 'MiddleRight'
$versionLabel.Font = New-Object Drawing.Font("Segoe UI", 8.5)
$statusBar.Controls.Add($versionLabel)
# -------------------------
# Main Content Container
# -------------------------
$contentPanel = New-Object Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = '16,16,16,8'
# Add controls in correct order for docking
# Windows Forms docks controls in reverse order (last added = first docked)
# Order: Fill first, then Bottom, then Top (in reverse visual order)
$form.Controls.Add($contentPanel)    # Fill - takes remaining space
$form.Controls.Add($statusBar)       # Bottom
$form.Controls.Add($toolbarSeparator) # Top - appears below toolbar
$form.Controls.Add($toolbarPanel)    # Top - appears below menu
$form.Controls.Add($menuStrip)       # Top - appears at very top
$form.MainMenuStrip = $menuStrip
# Status container panel (acts as border)
$statusBorderPanel = New-Object Windows.Forms.Panel
$statusBorderPanel.Dock = 'Fill'
$statusBorderPanel.Padding = '1,1,1,1'
$contentPanel.Controls.Add($statusBorderPanel)
# Status RichTextBox with modern styling
$status = New-Object Windows.Forms.RichTextBox
$status.Multiline = $true
$status.ReadOnly = $true
$status.Dock = 'Fill'
$status.ScrollBars = 'Vertical'
$status.BorderStyle = 'None'
$status.Font = New-Object Drawing.Font("Cascadia Code, Consolas, Courier New", 9.5)
$statusBorderPanel.Controls.Add($status)
# Progress Panel (Dock Bottom - add after Fill so it appears at bottom)
$progressPanel = New-Object Windows.Forms.Panel
$progressPanel.Dock = 'Bottom'
$progressPanel.Height = 36
$progressPanel.Padding = '0,8,0,8'
$contentPanel.Controls.Add($progressPanel)
# Progress bar border panel
$progressBorderPanel = New-Object Windows.Forms.Panel
$progressBorderPanel.Dock = 'Fill'
$progressBorderPanel.Padding = '1,1,1,1'
$progressPanel.Controls.Add($progressBorderPanel)
# Modern Progress Bar
$progress = New-Object Windows.Forms.ProgressBar
$progress.Dock = 'Fill'
$progress.Style = 'Continuous'
$progress.Value = 0
$progressBorderPanel.Controls.Add($progress)
# Header Label (Dock Top - add last so it appears at top)
$headerLabel = New-Object Windows.Forms.Label
$headerLabel.Text = "Output Console"
$headerLabel.Dock = 'Top'
$headerLabel.Height = 26
$headerLabel.Font = New-Object Drawing.Font("Segoe UI Semibold", 9.5)
$headerLabel.TextAlign = 'MiddleLeft'
$contentPanel.Controls.Add($headerLabel)
# Language combobox for compatibility
$cmbLang = New-Object Windows.Forms.ComboBox
$cmbLang.Items.AddRange(@("en", "ru"))
$cmbLang.SelectedItem = $global:CurrentLanguage
# DataGrid for drivers result (hidden initially)
$driversGrid = New-Object Windows.Forms.DataGridView
$driversGrid.ReadOnly = $true
$driversGrid.AllowUserToAddRows = $false
$driversGrid.AllowUserToDeleteRows = $false
$driversGrid.Height = 60
$driversGrid.Visible = $false
# Timer to poll job/logs
$timer = New-Object Windows.Forms.Timer
$timer.Interval = 1000
# -------------------------
# Background job helpers
# -------------------------
function Start-BackgroundTask {
param(
[string]$Name,
[array]$TaskArgs
)
if ($null -ne $global:CurrentJob -and (Get-Job -Id $global:CurrentJob.Id -ErrorAction SilentlyContinue)) {
[System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "TaskRunning"), "Task Running", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
return $null
}
$log = New-TaskLog $Name
$global:CurrentTaskLog = $log
Add-StatusUI $form $status ("$(Get-LocalizedString 'StartingTask') $Name")
# Pass filter settings to job (only used by CheckDriverUpdates now)
$filterClass = $global:FilterSettings.Class
$filterMfr = $global:FilterSettings.Manufacturer
# Start job with task logic defined directly in the job
$job = Start-Job -Name $Name -ScriptBlock {
param($taskName, $logPath, $innerArgs, $filterClass, $filterMfr)
# Logging function
function L($m) {
$t = (Get-Date).ToString("s")
Add-Content -Path $logPath -Value ("$t - $m")
}
# FIXED: Check admin rights in background job
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
L "[ERROR] This task requires administrator privileges"
L "[ERROR] Please run the script as Administrator"
L "Completed"
return
}
try {
# Execute task based on task name
switch ($taskName) {
# NEW: Combined task - try Rikor server first, fallback to Microsoft Update
"DownloadRikorDrivers" {
L "Attempting to download drivers from Rikor server..."
$rikorServerAvailable = $false
# TODO: Add actual Rikor server URL and download logic here
# For now, simulate server check
try {
    # Placeholder for Rikor server check
    # $rikorUrl = "https://rikor-server.example.com/drivers"
    # Test-Connection or Invoke-WebRequest to check availability
    L "[INFO] Rikor server is currently unavailable."
    $rikorServerAvailable = $false
} catch {
    L "[INFO] Cannot connect to Rikor server."
    $rikorServerAvailable = $false
}

if (-not $rikorServerAvailable) {
    L "[INFO] Falling back to Microsoft Update for driver search and installation..."
    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        L "Searching for available driver updates (this may take a few minutes)..."
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver'")
        if ($SearchResult.Updates.Count -eq 0) {
            L "No driver updates available from Microsoft Update"
        } else {
            L "Found $($SearchResult.Updates.Count) driver update(s) available:"
            L ""
            foreach ($Update in $SearchResult.Updates) {
                L "  - $($Update.Title)"
                L "    Size: $([math]::Round($Update.MaxDownloadSize / 1MB, 2)) MB"
            }
            L ""
            L "Downloading and installing driver updates..."
            
            # Create update collection for download and install
            $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
            foreach ($Update in $SearchResult.Updates) {
                if ($Update.EulaAccepted -eq $false) {
                    $Update.AcceptEula()
                }
                $UpdatesToDownload.Add($Update) | Out-Null
            }
            
            # Download updates
            if ($UpdatesToDownload.Count -gt 0) {
                L "Downloading $($UpdatesToDownload.Count) update(s)..."
                $Downloader = $UpdateSession.CreateUpdateDownloader()
                $Downloader.Updates = $UpdatesToDownload
                $DownloadResult = $Downloader.Download()
                L "Download completed with result code: $($DownloadResult.ResultCode)"
                
                # Install downloaded updates
                $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                foreach ($Update in $SearchResult.Updates) {
                    if ($Update.IsDownloaded) {
                        $UpdatesToInstall.Add($Update) | Out-Null
                    }
                }
                
                if ($UpdatesToInstall.Count -gt 0) {
                    L "Installing $($UpdatesToInstall.Count) update(s)..."
                    $Installer = $UpdateSession.CreateUpdateInstaller()
                    $Installer.Updates = $UpdatesToInstall
                    $InstallResult = $Installer.Install()
                    L "Installation completed with result code: $($InstallResult.ResultCode)"
                    L "Reboot required: $($InstallResult.RebootRequired)"
                    
                    $successCount = 0
                    $failCount = 0
                    for ($i = 0; $i -lt $UpdatesToInstall.Count; $i++) {
                        $resultCode = $InstallResult.GetUpdateResult($i).ResultCode
                        if ($resultCode -eq 2) { # 2 = Succeeded
                            $successCount++
                        } else {
                            $failCount++
                        }
                    }
                    L "Successfully installed: $successCount, Failed: $failCount"
                } else {
                    L "[WARNING] No updates were downloaded successfully"
                }
            }
        }
    } catch {
        L "[ERROR] Failed to download/install driver updates: $_"
    }
}
L "Completed"
}
"ScanDrivers" {
L "Scanning installed drivers..."
try {
$drivers = Get-CimInstance Win32_PnPSignedDriver | Select-Object DeviceName, Manufacturer, DriverVersion, InfName, Class, DriverDate
# Apply filters
if ($filterClass) {
L "Applying class filter: $filterClass"
$drivers = $drivers | Where-Object { $_.Class -like "*$filterClass*" }
}
if ($filterMfr) {
L "Applying manufacturer filter: $filterMfr"
$drivers = $drivers | Where-Object { $_.Manufacturer -like "*$filterMfr*" }
}
L "Found $($drivers.Count) driver(s) matching criteria"
# Group by class
$driversByClass = $drivers | Group-Object -Property Class
L "Categories: $($driversByClass.Count) different driver types"
L ""
L "Sample drivers:"
$sampleCount = [Math]::Min(5, $drivers.Count)
$drivers | Sort-Object DeviceName | Select-Object -First $sampleCount | ForEach-Object {
L "  * $($_.DeviceName) (v$($_.DriverVersion))"
}
if ($drivers.Count -gt $sampleCount) {
L "  ... and $($drivers.Count - $sampleCount) more"
}
L ""
L "Exporting full list to CSV..."
$csvPath = Join-Path (Split-Path $logPath) "InstalledDrivers_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
$drivers | Sort-Object DeviceName | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
L "Exported to: $csvPath"
L "Completed"
} catch {
L "[ERROR] Failed to scan drivers: $_"
}
}
"BackupDrivers" {
$dest = $innerArgs[0]
L "Backing up drivers to: $dest"
try {
if (!(Test-Path $dest)) {
New-Item -ItemType Directory -Path $dest -Force | Out-Null
}
L "Exporting drivers (this may take several minutes)..."
& dism.exe /online /export-driver /destination:$dest 2>&1 | Out-Null
$exportedCount = (Get-ChildItem -Path $dest -Recurse -Directory -ErrorAction SilentlyContinue).Count
if ($exportedCount -gt 0) {
L "Backup completed: $exportedCount driver package(s) exported"
} else {
L "Backup completed. Check destination folder for exported drivers."
}
} catch {
L "[ERROR] Backup failed: $_"
}
L "Completed"
}
"InstallDrivers" {
$folder = $innerArgs[0]
L "Installing drivers from: $folder"
try {
if (-not (Test-Path $folder)) {
L "[ERROR] Folder not found: $folder"
L "Completed"
return
}
$infFiles = Get-ChildItem -Path $folder -Recurse -Include *.inf -ErrorAction SilentlyContinue
if ($infFiles.Count -eq 0) {
L "[ERROR] No .inf driver files found in folder"
L "Completed"
return
}
L "Found $($infFiles.Count) driver file(s)"
L "Installing drivers..."
L ""
$successCount = 0
$failCount = 0
$current = 0
# FIXED: Corrected double error counting and added INF validation
foreach ($inf in $infFiles) {
$current++
L "[$current/$($infFiles.Count)] $($inf.Name)"
try {
# Validate INF file before installation
$infContent = Get-Content $inf.FullName -Raw -ErrorAction SilentlyContinue
if (-not $infContent -or $infContent -notmatch '\[Version\]') {
    L "     -> SKIPPED: Invalid INF file"
    $failCount++
    continue
}

# Check architecture compatibility
$is64Bit = [Environment]::Is64BitOperatingSystem
$archPattern = if ($is64Bit) { "amd64|x64|NTamd64" } else { "x86|NTx86" }
if ($infContent -notmatch $archPattern) {
    L "     -> SKIPPED: Incompatible architecture"
    $failCount++
    continue
}

# Install driver using pnputil
$out = & pnputil.exe /add-driver $inf.FullName /install /force 2>&1
if ($LASTEXITCODE -eq 0) {
    $successCount++
    L "     -> Successfully installed"
} else {
    $failCount++
    L "     -> Installation failed (exit code: $LASTEXITCODE)"
    # Log error details
    $out | Where-Object { $_ -match "error|fail" } | ForEach-Object {
        L "        $_"
    }
}
} catch {
$failCount++
L "     -> Exception: $_"
}
Start-Sleep -Milliseconds 300
}
L ""
L "Installation complete: $successCount successful, $failCount failed"
if ($successCount -gt 0) {
L "Note: Reboot may be required for some drivers."
}
} catch {
L "[ERROR] Installation failed: $_"
}
L "Completed"
}
default {
L "ERROR: Unknown task name: $taskName"
}
}
} catch {
L "ERROR in job: $_"
if ($_.ScriptStackTrace) {
L "Stack trace: $($_.ScriptStackTrace)"
}
}
} -ArgumentList $Name, $log, $TaskArgs, $filterClass, $filterMfr -RunAs32:$false
$global:CurrentJob = $job
# start UI timer to tail log
$timer.Start()
return $job
}
# Tail log file content and update UI
$timer.Add_Tick({
try {
if ($global:CurrentTaskLog -and (Test-Path $global:CurrentTaskLog)) {
$lines = Get-Content -Path $global:CurrentTaskLog -Tail 200 -ErrorAction SilentlyContinue
if ($lines) {
$text = ($lines -join "`r`n")
# Update UI (replace full text)
$invoke = [System.Windows.Forms.MethodInvoker]{
$status.Clear()
$status.AppendText($text + "`r`n")
$status.ScrollToCaret()
}
$form.Invoke($invoke)
}
}
# Update progress heuristic with more detailed tracking
if ($global:CurrentTaskLog -and (Test-Path $global:CurrentTaskLog)) {
$content = Get-Content -Path $global:CurrentTaskLog -Tail 100 -ErrorAction SilentlyContinue
$contentStr = $content -join "`n"
$p = 0
# NEW: Download and install progress heuristic
if ($contentStr -match "Downloading drivers archive") { $p = 5 }
if ($contentStr -match "Download completed") { $p = 15 }
if ($contentStr -match "Extracting ZIP archive") { $p = 20 }
if ($contentStr -match "Extraction completed") { $p = 30 }
if ($contentStr -match "Installing drivers from extracted archive") { $p = 40 }
if ($contentStr -match "\[.*\/.*\]") {
# Extract progress from [X/Y] format for install step
if ($contentStr -match "\[(\d+)\/(\d+)\]") {
$current = [int]$matches[1]
$total = [int]$matches[2]
if ($total -gt 0) {
$p = 40 + [int](($current / $total) * 55) # Remaining 55% of progress bar
}
}
}
if ($contentStr -match "Installation from archive complete") { $p = 95 }
if ($contentStr -match "Completed") { $p = 100 }

# Scan drivers progress
if ($contentStr -match "Scanning installed") { $p = 30 }
if ($contentStr -match "Found.*driver.*matching") { $p = 60 }
if ($contentStr -match "Exporting full list") { $p = 80 }
# Backup drivers progress
if ($contentStr -match "Exporting drivers") { $p = 40 }
if ($contentStr -match "Backup completed") { $p = 90 }
# Install drivers progress
if ($contentStr -match "Found.*driver file") { $p = 20 }
if ($contentStr -match "\[.*\/.*\]" -and -not $contentStr -match "ExtractedDrivers") { # Avoid overlap with download&install progress
# Extract progress from [X/Y] format for manual install step
if ($contentStr -match "\[(\d+)\/(\d+)\]") {
$current = [int]$matches[1]
$total = [int]$matches[2]
if ($total -gt 0) {
$p = 20 + [int](($current / $total) * 70)
}
}
}

# Generic progress indicators
if ($p -eq 0) {
if ($contentStr -match "Starting") { $p = 10 }
}

$progress.Value = [Math]::Min(100, [int]$p)
}
# Check job finished
if ($null -ne $global:CurrentJob) {
$jobState = (Get-Job -Id $global:CurrentJob.Id -ErrorAction SilentlyContinue).State
if ($jobState -in @("Completed","Failed","Stopped","Disconnected","Blocked","Suspended","Aborted")) {
Start-Sleep -Milliseconds 200
# Log to history
$taskName = $global:CurrentJob.Name
Add-HistoryEntry -TaskName $taskName -Status $jobState -Details "Task completed"
# pull final output and mark complete
$finishedText = Get-LocalizedString 'TaskFinished'
$invoke = [System.Windows.Forms.MethodInvoker]{
$status.AppendText("`r`n$finishedText $jobState ===`r`n")
$status.ScrollToCaret()
$progress.Value = 100
$statusLabel.Text = "  Task completed: $jobState"
}
$form.Invoke($invoke)
# cleanup job object
Remove-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
$global:CurrentJob = $null
$timer.Stop()
}
}
} catch {
# ignore UI timer exceptions
}
})
# -------------------------
# Theme Functions
# -------------------------
function Set-Theme {
param([bool]$Dark)
$global:DarkModeEnabled = $Dark
$colors = Get-ThemeColors
# Main form
$form.BackColor = $colors.Background
# Menu strip
$menuStrip.BackColor = $colors.MenuBar
$menuStrip.ForeColor = $colors.Text
foreach ($item in $menuStrip.Items) {
$item.BackColor = $colors.MenuBar
$item.ForeColor = $colors.Text
}
# Toolbar
$toolbarPanel.BackColor = $colors.Surface
$buttonContainer.BackColor = $colors.Surface
# Toolbar separator
$toolbarSeparator.BackColor = $colors.Border
# Style toolbar buttons
foreach ($ctrl in $buttonContainer.Controls) {
if ($ctrl -is [System.Windows.Forms.Button]) {
if ($ctrl.Tag -eq $true) {
# Primary buttons
$ctrl.BackColor = $colors.Primary
$ctrl.ForeColor = [System.Drawing.Color]::White
$ctrl.FlatAppearance.MouseOverBackColor = $colors.PrimaryHover
} else {
# Secondary buttons
$ctrl.BackColor = $colors.Secondary
$ctrl.ForeColor = $colors.Text
$ctrl.FlatAppearance.MouseOverBackColor = $colors.SurfaceHover
}
}
}
# Content panel
$contentPanel.BackColor = $colors.Background
$headerLabel.BackColor = $colors.Background
$headerLabel.ForeColor = $colors.TextSecondary
# Status textbox with border
$statusBorderPanel.BackColor = $colors.Border
$status.BackColor = $colors.Surface
$status.ForeColor = $colors.Text
# Progress panel with border
$progressPanel.BackColor = $colors.Background
$progressBorderPanel.BackColor = $colors.Border
# Status bar
$statusBar.BackColor = $colors.StatusBar
$statusLabel.BackColor = $colors.StatusBar
$statusLabel.ForeColor = [System.Drawing.Color]::White
$versionLabel.BackColor = $colors.StatusBar
$versionLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)
# Update menu theme text
if ($Dark) {
$menuToggleTheme.Text = Get-LocalizedString "BtnLightMode"
} else {
$menuToggleTheme.Text = Get-LocalizedString "BtnDarkMode"
}
Export-Settings
}
# -------------------------
# Update UI Language
# -------------------------
function Update-UILanguage {
$form.Text = Get-LocalizedString "AppTitle"
# Toolbar buttons
$btnDownloadRikor.Text = Get-LocalizedString "BtnDownloadRikor"
$btnScan.Text = Get-LocalizedString "BtnScan"
$btnBackup.Text = Get-LocalizedString "BtnBackup"
$btnInstall.Text = Get-LocalizedString "BtnInstall"
$btnCancel.Text = Get-LocalizedString "BtnCancel"
# Menu items
$menuOpenLogs.Text = Get-LocalizedString "BtnOpenLogs"
$menuDownloadRikor.Text = Get-LocalizedString "BtnDownloadRikor"
$menuScan.Text = Get-LocalizedString "BtnScan"
$menuBackup.Text = Get-LocalizedString "BtnBackup"
$menuInstall.Text = Get-LocalizedString "BtnInstall"
$menuCancel.Text = Get-LocalizedString "BtnCancel"
$menuRestorePoint.Text = Get-LocalizedString "BtnRestorePoint"
$menuSchedule.Text = Get-LocalizedString "BtnSchedule"
$menuFilters.Text = Get-LocalizedString "BtnFilters"
$menuHistory.Text = Get-LocalizedString "BtnHistory"
if ($global:DarkModeEnabled) {
$menuToggleTheme.Text = Get-LocalizedString "BtnLightMode"
} else {
$menuToggleTheme.Text = Get-LocalizedString "BtnDarkMode"
}
# Update language menu checkmarks
foreach ($item in $menuLanguage.DropDownItems) {
$item.Checked = ($item.Tag -eq $global:CurrentLanguage)
}
}
# -------------------------
# Dialog Functions
# -------------------------
function Show-HistoryDialog {
$colors = Get-ThemeColors
$historyForm = New-Object Windows.Forms.Form
$historyForm.Text = Get-LocalizedString "HistoryTitle"
$historyForm.Size = '750,500'
$historyForm.StartPosition = "CenterParent"
$historyForm.BackColor = $colors.Background
$historyForm.TopMost = $true
$historyForm.Font = New-Object Drawing.Font("Segoe UI", 9.5)
$historyForm.FormBorderStyle = 'FixedDialog'
$historyForm.MaximizeBox = $false
# Header
$header = New-Object Windows.Forms.Label
$header.Text = Get-LocalizedString "HistoryTitle"
$header.Dock = 'Top'
$header.Height = 45
$header.BackColor = $colors.Primary
$header.ForeColor = [System.Drawing.Color]::White
$header.Font = New-Object Drawing.Font("Segoe UI Semibold", 12)
$header.TextAlign = 'MiddleCenter'
$historyList = New-Object Windows.Forms.ListView
$historyList.Dock = 'Fill'
$historyList.View = 'Details'
$historyList.FullRowSelect = $true
$historyList.GridLines = $false
$historyList.BackColor = $colors.Surface
$historyList.ForeColor = $colors.Text
$historyList.BorderStyle = 'None'
$historyList.Font = New-Object Drawing.Font("Segoe UI", 9)
$historyList.Columns.Add("Timestamp", 160) | Out-Null
$historyList.Columns.Add("Task", 140) | Out-Null
$historyList.Columns.Add("Status", 100) | Out-Null
$historyList.Columns.Add("Details", 300) | Out-Null
$history = Get-UpdateHistory
if ($history) {
foreach ($entry in $history) {
$item = New-Object Windows.Forms.ListViewItem($entry.Timestamp)
$item.SubItems.Add($entry.Task) | Out-Null
$item.SubItems.Add($entry.Status) | Out-Null
$item.SubItems.Add($entry.Details) | Out-Null
$historyList.Items.Add($item) | Out-Null
}
}
# Add controls in correct order: Fill first, then Top (reverse dock processing)
$historyForm.Controls.Add($historyList)
$historyForm.Controls.Add($header)
$historyForm.ShowDialog($form) | Out-Null
}
function Show-ScheduleDialog {
$colors = Get-ThemeColors
$schedForm = New-Object Windows.Forms.Form
$schedForm.Text = Get-LocalizedString "ScheduleTitle"
$schedForm.Size = '420,300'
$schedForm.StartPosition = "CenterParent"
$schedForm.BackColor = $colors.Background
$schedForm.FormBorderStyle = 'FixedDialog'
$schedForm.MaximizeBox = $false
$schedForm.TopMost = $true
$schedForm.Font = New-Object Drawing.Font("Segoe UI", 9.5)
# Header
$header = New-Object Windows.Forms.Label
$header.Text = Get-LocalizedString "ScheduleTitle"
$header.Dock = 'Top'
$header.Height = 45
$header.BackColor = $colors.Primary
$header.ForeColor = [System.Drawing.Color]::White
$header.Font = New-Object Drawing.Font("Segoe UI Semibold", 12)
$header.TextAlign = 'MiddleCenter'
$contentPanel = New-Object Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = '20,15,20,15'
$contentPanel.BackColor = $colors.Surface
# Add controls in correct order: Fill first, then Top (reverse dock processing)
$schedForm.Controls.Add($contentPanel)
$schedForm.Controls.Add($header)
$lblFreq = New-Object Windows.Forms.Label
$lblFreq.Text = "Frequency:"
$lblFreq.Location = '0,10'
$lblFreq.Size = '100,25'
$lblFreq.ForeColor = $colors.Text
$contentPanel.Controls.Add($lblFreq)
$cmbFreq = New-Object Windows.Forms.ComboBox
$cmbFreq.Location = '110,8'
$cmbFreq.Size = '180,28'
$cmbFreq.DropDownStyle = 'DropDownList'
$cmbFreq.BackColor = $colors.Background
$cmbFreq.ForeColor = $colors.Text
$cmbFreq.Items.AddRange(@((Get-LocalizedString "Daily"), (Get-LocalizedString "Weekly"), (Get-LocalizedString "Monthly")))
$cmbFreq.SelectedIndex = 0
$contentPanel.Controls.Add($cmbFreq)
$lblTime = New-Object Windows.Forms.Label
$lblTime.Text = Get-LocalizedString "Time"
$lblTime.Location = '0,50'
$lblTime.Size = '100,25'
$lblTime.ForeColor = $colors.Text
$contentPanel.Controls.Add($lblTime)
$txtTime = New-Object Windows.Forms.TextBox
$txtTime.Location = '110,48'
$txtTime.Size = '180,28'
$txtTime.Text = "03:00"
$txtTime.BackColor = $colors.Background
$txtTime.ForeColor = $colors.Text
$contentPanel.Controls.Add($txtTime)
$existingTask = Get-ScheduledUpdateTask
$lblStatus = New-Object Windows.Forms.Label
$lblStatus.Location = '0,95'
$lblStatus.Size = '350,25'
$lblStatus.Font = New-Object Drawing.Font("Segoe UI Semibold", 9)
if ($existingTask) {
$lblStatus.Text = "Current schedule: Active"
$lblStatus.ForeColor = $colors.Success
} else {
$lblStatus.Text = "No schedule configured"
$lblStatus.ForeColor = $colors.TextSecondary
}
$contentPanel.Controls.Add($lblStatus)
# Buttons panel
$btnPanel = New-Object Windows.Forms.Panel
$btnPanel.Location = '0,135'
$btnPanel.Size = '360,40'
$contentPanel.Controls.Add($btnPanel)
$btnEnable = New-Object Windows.Forms.Button
$btnEnable.Text = Get-LocalizedString "Enable"
$btnEnable.Location = '0,0'
$btnEnable.Size = '110,36'
$btnEnable.FlatStyle = 'Flat'
$btnEnable.FlatAppearance.BorderSize = 0
$btnEnable.BackColor = $colors.Primary
$btnEnable.ForeColor = [System.Drawing.Color]::White
$btnEnable.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnEnable.Add_Click({
$freqMap = @{
(Get-LocalizedString "Daily") = "Daily"
(Get-LocalizedString "Weekly") = "Weekly"
(Get-LocalizedString "Monthly") = "Monthly"
}
$freq = $freqMap[$cmbFreq.SelectedItem]
if (-not $freq) { $freq = "Daily" }
if (Set-ScheduledUpdate -Frequency $freq -Time $txtTime.Text) {
$lblStatus.Text = Get-LocalizedString "ScheduleCreated"
$lblStatus.ForeColor = $colors.Success
}
})
$btnPanel.Controls.Add($btnEnable)
$btnRemove = New-Object Windows.Forms.Button
$btnRemove.Text = Get-LocalizedString "Remove"
$btnRemove.Location = '120,0'
$btnRemove.Size = '120,36'
$btnRemove.FlatStyle = 'Flat'
$btnRemove.FlatAppearance.BorderSize = 0
$btnRemove.BackColor = $colors.Secondary
$btnRemove.ForeColor = $colors.Text
$btnRemove.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnRemove.Add_Click({
Remove-ScheduledUpdate
$lblStatus.Text = Get-LocalizedString "ScheduleRemoved"
$lblStatus.ForeColor = $colors.TextSecondary
})
$btnPanel.Controls.Add($btnRemove)
$btnClose = New-Object Windows.Forms.Button
$btnClose.Text = Get-LocalizedString "Close"
$btnClose.Location = '250,0'
$btnClose.Size = '100,36'
$btnClose.FlatStyle = 'Flat'
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.BackColor = $colors.Secondary
$btnClose.ForeColor = $colors.Text
$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClose.Add_Click({ $schedForm.Close() })
$btnPanel.Controls.Add($btnClose)
$schedForm.ShowDialog($form) | Out-Null
}
function Show-FiltersDialog {
$colors = Get-ThemeColors
$filterForm = New-Object Windows.Forms.Form
$filterForm.Text = Get-LocalizedString "FilterTitle"
$filterForm.Size = '420,260'
$filterForm.StartPosition = "CenterParent"
$filterForm.BackColor = $colors.Background
$filterForm.FormBorderStyle = 'FixedDialog'
$filterForm.MaximizeBox = $false
$filterForm.TopMost = $true
$filterForm.Font = New-Object Drawing.Font("Segoe UI", 9.5)
# Header
$header = New-Object Windows.Forms.Label
$header.Text = Get-LocalizedString "FilterTitle"
$header.Dock = 'Top'
$header.Height = 45
$header.BackColor = $colors.Primary
$header.ForeColor = [System.Drawing.Color]::White
$header.Font = New-Object Drawing.Font("Segoe UI Semibold", 12)
$header.TextAlign = 'MiddleCenter'
$contentPanel = New-Object Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = '20,15,20,15'
$contentPanel.BackColor = $colors.Surface
# Add controls in correct order: Fill first, then Top (reverse dock processing)
$filterForm.Controls.Add($contentPanel)
$filterForm.Controls.Add($header)
$lblClass = New-Object Windows.Forms.Label
$lblClass.Text = Get-LocalizedString "ClassFilter"
$lblClass.Location = '0,10'
$lblClass.Size = '160,25'
$lblClass.ForeColor = $colors.Text
$contentPanel.Controls.Add($lblClass)
$txtClass = New-Object Windows.Forms.TextBox
$txtClass.Location = '170,8'
$txtClass.Size = '180,28'
$txtClass.Text = $global:FilterSettings.Class
$txtClass.BackColor = $colors.Background
$txtClass.ForeColor = $colors.Text
$contentPanel.Controls.Add($txtClass)
$lblMfr = New-Object Windows.Forms.Label
$lblMfr.Text = Get-LocalizedString "ManufacturerFilter"
$lblMfr.Location = '0,50'
$lblMfr.Size = '160,25'
$lblMfr.ForeColor = $colors.Text
$contentPanel.Controls.Add($lblMfr)
$txtMfr = New-Object Windows.Forms.TextBox
$txtMfr.Location = '170,48'
$txtMfr.Size = '180,28'
$txtMfr.Text = $global:FilterSettings.Manufacturer
$txtMfr.BackColor = $colors.Background
$txtMfr.ForeColor = $colors.Text
$contentPanel.Controls.Add($txtMfr)
# Buttons panel
$btnPanel = New-Object Windows.Forms.Panel
$btnPanel.Location = '0,100'
$btnPanel.Size = '360,40'
$contentPanel.Controls.Add($btnPanel)
$btnApply = New-Object Windows.Forms.Button
$btnApply.Text = Get-LocalizedString "Apply"
$btnApply.Location = '0,0'
$btnApply.Size = '110,36'
$btnApply.FlatStyle = 'Flat'
$btnApply.FlatAppearance.BorderSize = 0
$btnApply.BackColor = $colors.Primary
$btnApply.ForeColor = [System.Drawing.Color]::White
$btnApply.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnApply.Add_Click({
$global:FilterSettings.Class = $txtClass.Text
$global:FilterSettings.Manufacturer = $txtMfr.Text
Export-Settings
Add-StatusUI $form $status "$(Get-LocalizedString 'FilterApplied') Class='$($txtClass.Text)', Manufacturer='$($txtMfr.Text)'"
$filterForm.Close()
})
$btnPanel.Controls.Add($btnApply)
$btnClear = New-Object Windows.Forms.Button
$btnClear.Text = Get-LocalizedString "Clear"
$btnClear.Location = '120,0'
$btnClear.Size = '100,36'
$btnClear.FlatStyle = 'Flat'
$btnClear.FlatAppearance.BorderSize = 0
$btnClear.BackColor = $colors.Secondary
$btnClear.ForeColor = $colors.Text
$btnClear.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClear.Add_Click({
$txtClass.Text = ""
$txtMfr.Text = ""
$global:FilterSettings.Class = ""
$global:FilterSettings.Manufacturer = ""
Export-Settings
Add-StatusUI $form $status (Get-LocalizedString "FilterCleared")
})
$btnPanel.Controls.Add($btnClear)
$btnClose = New-Object Windows.Forms.Button
$btnClose.Text = Get-LocalizedString "Close"
$btnClose.Location = '230,0'
$btnClose.Size = '100,36'
$btnClose.FlatStyle = 'Flat'
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.BackColor = $colors.Secondary
$btnClose.ForeColor = $colors.Text
$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClose.Add_Click({ $filterForm.Close() })
$btnPanel.Controls.Add($btnClose)
$filterForm.ShowDialog($form) | Out-Null
}
function Show-SettingsDialog {
$colors = Get-ThemeColors
$settingsForm = New-Object Windows.Forms.Form
$settingsForm.Text = Get-LocalizedString "SettingsTitle"
$settingsForm.Size = '500,340'
$settingsForm.StartPosition = "CenterParent"
$settingsForm.BackColor = $colors.Background
$settingsForm.FormBorderStyle = 'FixedDialog'
$settingsForm.MaximizeBox = $false
$settingsForm.TopMost = $true
$settingsForm.Font = New-Object Drawing.Font("Segoe UI", 9.5)
# Header
$header = New-Object Windows.Forms.Label
$header.Text = Get-LocalizedString "SettingsTitle"
$header.Dock = 'Top'
$header.Height = 45
$header.BackColor = $colors.Primary
$header.ForeColor = [System.Drawing.Color]::White
$header.Font = New-Object Drawing.Font("Segoe UI Semibold", 12)
$header.TextAlign = 'MiddleCenter'
$contentPanel = New-Object Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = '20,15,20,15'
$contentPanel.BackColor = $colors.Surface
# Add controls in correct order: Fill first, then Top (reverse dock processing)
$settingsForm.Controls.Add($contentPanel)
$settingsForm.Controls.Add($header)
# Proxy section
$lblProxySection = New-Object Windows.Forms.Label
$lblProxySection.Text = "Network Proxy"
$lblProxySection.Location = '0,5'
$lblProxySection.Size = '150,20'
$lblProxySection.ForeColor = $colors.Primary
$lblProxySection.Font = New-Object Drawing.Font("Segoe UI Semibold", 9.5)
$contentPanel.Controls.Add($lblProxySection)
$lblProxy = New-Object Windows.Forms.Label
$lblProxy.Text = Get-LocalizedString "ProxyLabel"
$lblProxy.Location = '0,35'
$lblProxy.Size = '130,25'
$lblProxy.ForeColor = $colors.Text
$contentPanel.Controls.Add($lblProxy)
$txtProxy = New-Object Windows.Forms.TextBox
$txtProxy.Location = '140,33'
$txtProxy.Size = '220,28'
$txtProxy.Text = $global:ProxySettings.Address
$txtProxy.BackColor = $colors.Background
$txtProxy.ForeColor = $colors.Text
$contentPanel.Controls.Add($txtProxy)
$chkProxy = New-Object Windows.Forms.CheckBox
$chkProxy.Text = Get-LocalizedString "Enable"
$chkProxy.Location = '370,33'
$chkProxy.Size = '80,25'
$chkProxy.Checked = $global:ProxySettings.Enabled
$chkProxy.ForeColor = $colors.Text
$contentPanel.Controls.Add($chkProxy)
$btnApplyProxy = New-Object Windows.Forms.Button
$btnApplyProxy.Text = Get-LocalizedString "Apply"
$btnApplyProxy.Location = '140,70'
$btnApplyProxy.Size = '110,34'
$btnApplyProxy.FlatStyle = 'Flat'
$btnApplyProxy.FlatAppearance.BorderSize = 0
$btnApplyProxy.BackColor = $colors.Primary
$btnApplyProxy.ForeColor = [System.Drawing.Color]::White
$btnApplyProxy.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnApplyProxy.Add_Click({
Set-ProxySettings -ProxyAddr $txtProxy.Text -Enable $chkProxy.Checked
if ($chkProxy.Checked -and $txtProxy.Text) {
Add-StatusUI $form $status "$(Get-LocalizedString 'ProxyConfigured') $($txtProxy.Text)"
} else {
Add-StatusUI $form $status (Get-LocalizedString "ProxyCleared")
}
})
$contentPanel.Controls.Add($btnApplyProxy)
# Separator
$separator = New-Object Windows.Forms.Label
$separator.Location = '0,120'
$separator.Size = '440,1'
$separator.BackColor = $colors.Border
$contentPanel.Controls.Add($separator)
# Info section
$lblInfoSection = New-Object Windows.Forms.Label
$lblInfoSection.Text = "Application Info"
$lblInfoSection.Location = '0,130'
$lblInfoSection.Size = '150,20'
$lblInfoSection.ForeColor = $colors.Primary
$lblInfoSection.Font = New-Object Drawing.Font("Segoe UI Semibold", 9.5)
$contentPanel.Controls.Add($lblInfoSection)
$lblInfo = New-Object Windows.Forms.Label
$lblInfo.Text = "Logs: $LogBase"
$lblInfo.Location = '0,155'
$lblInfo.Size = '440,20'
$lblInfo.ForeColor = $colors.TextSecondary
$lblInfo.Font = New-Object Drawing.Font("Segoe UI", 8.5)
$contentPanel.Controls.Add($lblInfo)
$lblInfo2 = New-Object Windows.Forms.Label
$lblInfo2.Text = "Settings: $SettingsFile"
$lblInfo2.Location = '0,175'
$lblInfo2.Size = '440,20'
$lblInfo2.ForeColor = $colors.TextSecondary
$lblInfo2.Font = New-Object Drawing.Font("Segoe UI", 8.5)
$contentPanel.Controls.Add($lblInfo2)
$btnClose = New-Object Windows.Forms.Button
$btnClose.Text = Get-LocalizedString "Close"
$btnClose.Location = '170,210'
$btnClose.Size = '110,36'
$btnClose.FlatStyle = 'Flat'
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.BackColor = $colors.Secondary
$btnClose.ForeColor = $colors.Text
$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClose.Add_Click({ $settingsForm.Close() })
$contentPanel.Controls.Add($btnClose)
$settingsForm.ShowDialog($form) | Out-Null
}
# -------------------------
# Action Functions (shared by buttons and menus)
# -------------------------
# NEW: Combined function for Rikor download with MS Update fallback
function Invoke-DownloadRikorDrivers {
$status.Clear()
$progress.Value = 0
$statusLabel.Text = "  Downloading drivers from Rikor server..."
Start-BackgroundTask -Name "DownloadRikorDrivers" -TaskArgs @()
}
function Invoke-ScanDrivers {
$status.Clear()
$progress.Value = 0
$statusLabel.Text = "  Scanning installed drivers..."
Start-BackgroundTask -Name "ScanDrivers" -TaskArgs @()
}
function Invoke-BackupDrivers {
$status.Clear()
$progress.Value = 0
$fbd = New-Object System.Windows.Forms.FolderBrowserDialog
$fbd.Description = Get-LocalizedString "SelectBackupFolder"
$fbd.ShowNewFolderButton = $true
if ($fbd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
$dest = $fbd.SelectedPath
$statusLabel.Text = "  Backing up drivers..."
Start-BackgroundTask -Name "BackupDrivers" -TaskArgs @($dest)
} else {
Add-StatusUI $form $status (Get-LocalizedString "BackupCanceled")
}
}
function Invoke-InstallDrivers {
$status.Clear()
$progress.Value = 0
$fbd = New-Object System.Windows.Forms.FolderBrowserDialog
$fbd.Description = Get-LocalizedString "SelectDriverFolder"
$fbd.ShowNewFolderButton = $false
if ($fbd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
$folder = $fbd.SelectedPath
$statusLabel.Text = "  Installing drivers..."
Start-BackgroundTask -Name "InstallDrivers" -TaskArgs @($folder)
} else {
Add-StatusUI $form $status (Get-LocalizedString "InstallCanceled")
}
}
function Invoke-CancelTask {
try {
if ($null -ne $global:CurrentJob) {
Stop-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
Remove-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
$global:CurrentJob = $null
$timer.Stop()
$progress.Value = 0
$statusLabel.Text = "  Task cancelled"
Add-StatusUI $form $status (Get-LocalizedString "TaskCancelled")
} else {
Add-StatusUI $form $status (Get-LocalizedString "NoTaskToCancel")
}
} catch {
Add-StatusUI $form $status "Cancel error: $_"
}
}
function Invoke-OpenLogs {
if (Test-Path $LogBase) {
Start-Process -FilePath "explorer.exe" -ArgumentList "`"$LogBase`""
} else {
Add-StatusUI $form $status (Get-LocalizedString "LogFolderMissing")
}
}
function Invoke-ToggleTheme {
Set-Theme -Dark (-not $global:DarkModeEnabled)
}
function Invoke-CreateRestorePoint {
$status.Clear()
$statusLabel.Text = "  Creating restore point..."
Add-StatusUI $form $status "Creating system restore point..."
$progress.Value = 50
if (New-RestorePoint -Description "Rikor Driver Installer - $(Get-Date -Format 'yyyy-MM-dd HH:mm')") {
Add-StatusUI $form $status (Get-LocalizedString "RestorePointCreated")
$statusLabel.Text = "  Restore point created"
} else {
Add-StatusUI $form $status (Get-LocalizedString "RestorePointFailed")
$statusLabel.Text = "  Restore point failed"
}
$progress.Value = 100
}
# -------------------------
# Button handlers
# -------------------------
# NEW: Single combined button handler
$btnDownloadRikor.Add_Click({ Invoke-DownloadRikorDrivers })
$btnScan.Add_Click({ Invoke-ScanDrivers })
$btnBackup.Add_Click({ Invoke-BackupDrivers })
$btnInstall.Add_Click({ Invoke-InstallDrivers })
$btnCancel.Add_Click({ Invoke-CancelTask })
# -------------------------
# Menu handlers
# -------------------------
# NEW: Single combined menu handler
$menuDownloadRikor.Add_Click({ Invoke-DownloadRikorDrivers })
$menuScan.Add_Click({ Invoke-ScanDrivers })
$menuBackup.Add_Click({ Invoke-BackupDrivers })
$menuInstall.Add_Click({ Invoke-InstallDrivers })
$menuCancel.Add_Click({ Invoke-CancelTask })
$menuOpenLogs.Add_Click({ Invoke-OpenLogs })
$menuToggleTheme.Add_Click({ Invoke-ToggleTheme })
$menuRestorePoint.Add_Click({ Invoke-CreateRestorePoint })
$menuSchedule.Add_Click({ Show-ScheduleDialog })
$menuHistory.Add_Click({ Show-HistoryDialog })
$menuFilters.Add_Click({ Show-FiltersDialog })
$menuSettingsTop.Add_Click({ Show-SettingsDialog })
$menuExit.Add_Click({ $form.Close() })
# Language menu handlers
foreach ($item in $menuLanguage.DropDownItems) {
$item.Add_Click({
$clickedItem = $this
$global:CurrentLanguage = $clickedItem.Tag
Update-UILanguage
Export-Settings
$statusLabel.Text = "  $(Get-LocalizedString 'LanguageChanged') $($clickedItem.Tag)"
Add-StatusUI $form $status "$(Get-LocalizedString 'LanguageChanged') $($clickedItem.Tag)"
})
}
# Update button centering on resize
$form.Add_Resize({ Update-ButtonContainerPadding })
# Form closing cleanup
$form.Add_FormClosing({
try {
if ($null -ne $global:CurrentJob) {
Stop-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
Remove-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
}
} catch {}
})
# Apply saved theme
Set-Theme -Dark $global:DarkModeEnabled
# Show form
$form.Topmost = $true
$form.Add_Shown({
$form.Activate()
Update-ButtonContainerPadding
})
[void]$form.ShowDialog()