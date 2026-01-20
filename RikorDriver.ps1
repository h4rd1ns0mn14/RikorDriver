# Enhanced version with proper error handling, internet connectivity checks, and dual language support
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
# Multi-Language Support (Only English and Russian)
# -------------------------
$global:Languages = @{
    "en" = @{
        AppTitle = "Rikor Driver Installer"
        BtnWU = "Check Windows Updates"
        BtnCheckUpdates = "Check Microsoft Updates"
        BtnRikorUpdate = "Install Rikor Drivers" # NEW: Button for Rikor-specific drivers
        BtnDownloadAndInstall = "Download & Install Rikor Drivers"
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
        InternetCheckFailed = "Internet connection check failed. Some features may not work properly."
        AdminCheckFailed = "Administrator privileges check failed."
        UpdatesAvailable = "Updates available:"
        NoUpdatesAvailable = "No updates available."
        RikorServerUnavailable = "Rikor Update server is currently unavailable."
        CheckingForUpdates = "Checking for updates..."
        InstallingUpdates = "Installing updates..."
        UpdateSuccess = "Updates installed successfully!"
        UpdateFailed = "Update installation failed."
    }
    "ru" = @{
        AppTitle = "Установщик драйверов Rikor"
        BtnWU = "Проверить обновления Windows"
        BtnCheckUpdates = "Проверить обновления Microsoft"
        BtnRikorUpdate = "Установить драйверы Rikor" # NEW: Button for Rikor-specific drivers
        BtnDownloadAndInstall = "Загрузить и установить драйверы Rikor"
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
        InternetCheckFailed = "Проверка подключения к Интернету не удалась. Некоторые функции могут работать некорректно."
        AdminCheckFailed = "Проверка прав администратора не удалась."
        UpdatesAvailable = "Доступны обновления:"
        NoUpdatesAvailable = "Обновления отсутствуют."
        RikorServerUnavailable = "Сервер обновлений Rikor временно недоступен."
        CheckingForUpdates = "Проверка наличия обновлений..."
        InstallingUpdates = "Установка обновлений..."
        UpdateSuccess = "Обновления успешно установлены!"
        UpdateFailed = "Ошибка при установке обновлений."
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
# Enhanced Error Handling and Checks
# -------------------------

# Check for internet connectivity
function Test-InternetConnection {
    try {
        # Try to connect to Microsoft's website
        $response = Invoke-WebRequest -Uri "https://www.microsoft.com" -UseBasicParsing -TimeoutSec 10
        return $response.StatusCode -eq 200
    } catch {
        return $false
    }
}

# Check for admin privileges
function Test-AdminPrivileges {
    try {
        return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

# Check for Windows Updates availability
function Test-WindowsUpdates {
    try {
        # Check if Windows Update service is available
        $service = Get-Service -Name wuauserv -ErrorAction Stop
        return $service.Status -eq "Running"
    } catch {
        return $false
    }
}

# Check for Microsoft Update availability
function Test-MicrosoftUpdates {
    try {
        # Attempt to create a session with Windows Update
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        return $true
    } catch {
        return $false
    }
}

# Check for Rikor Update availability (stub implementation)
function Test-RikorUpdates {
    # This is a placeholder - implement actual Rikor update server check
    # Currently returns false as server is unavailable as mentioned in requirements
    return $false
}

# Enhanced admin privilege assertion with error handling
function Assert-AdminPrivilege {
    if (-not (Test-AdminPrivileges)) {
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "PermissionError"), "Permission Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        exit 1
    }
}

# Perform initial checks
Assert-AdminPrivilege

# Check internet connection
if (-not (Test-InternetConnection)) {
    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "InternetCheckFailed"), "Internet Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

# -------------------------
# Globals and paths
# -------------------------
$AppTitle = Get-LocalizedString "AppTitle"
$LogBase = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller"
$HistoryFile = Join-Path $LogBase "UpdateHistory.json"
$SettingsFile = Join-Path $LogBase "Settings.json"

if (!(Test-Path $LogBase)) { 
    New-Item -ItemType Directory -Path $LogBase -Force | Out-Null 
}

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
        } catch {
            Write-Warning "Failed to import settings: $_"
        }
    }
}

function Export-Settings {
    $settings = @{
        Language = $global:CurrentLanguage
        Proxy = $global:ProxySettings
        Filters = $global:FilterSettings
        DarkMode = $global:DarkModeEnabled
    }
    try {
        $settings | ConvertTo-Json -Depth 3 | Set-Content -Path $SettingsFile -Encoding UTF8
    } catch {
        Write-Warning "Failed to export settings: $_"
    }
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
                }
            }
        } catch {
            Write-Warning "Failed to parse history file: $_"
        }
    }
    
    $entry = @{
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Task = $TaskName
        Status = $Status
        Details = $Details
    }
    
    $history = @($entry) + $history
    $history = $history | Select-Object -First 100
    
    try {
        $history | ConvertTo-Json -Depth 4 | Set-Content -Path $HistoryFile -Encoding UTF8
    } catch {
        Write-Warning "Failed to write history file: $_"
    }
}

function Get-UpdateHistory {
    if (Test-Path $HistoryFile) {
        try {
            $content = Get-Content -Path $HistoryFile -Raw
            if ($content) {
                $parsed = $content | ConvertFrom-Json
                if ($parsed -is [array]) {
                    return $parsed
                }
            }
        } catch {
            Write-Warning "Failed to get history: $_"
        }
    }
    return @()
}

# -------------------------
# Task Logging
# -------------------------
function New-TaskLog([string]$TaskName) {
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $logFile = Join-Path $LogBase "$TaskName`_$timestamp.log"
    # Create log file with header
    $header = "# Log for task: $TaskName`n# Started: $(Get-Date)`n# User: $env:USERNAME`n# Computer: $env:COMPUTERNAME`n"
    Add-Content -Path $logFile -Value $header -Encoding UTF8
    return $logFile
}

function Add-StatusUI($form, $statusCtrl, $msg) {
    $now = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $fullMsg = "[$now] $msg"
    $invoke = [System.Windows.Forms.MethodInvoker]{
        $statusCtrl.AppendText("$fullMsg`r`n")
        $statusCtrl.ScrollToCaret()
    }
    $form.Invoke($invoke)
}

# -------------------------
# Theme Functions
# -------------------------
function Get-ThemeColors {
    if ($global:DarkModeEnabled) {
        return @{
            Background = [System.Drawing.Color]::FromArgb(30, 30, 30)
            Surface = [System.Drawing.Color]::FromArgb(45, 45, 45)
            Text = [System.Drawing.Color]::White
            TextSecondary = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)
            Border = [System.Drawing.Color]::FromArgb(65, 65, 65)
            Primary = [System.Drawing.Color]::FromArgb(0, 120, 215)  # Blue color instead of purple
            PrimaryHover = [System.Drawing.Color]::FromArgb(0, 100, 190)
            Secondary = [System.Drawing.Color]::FromArgb(65, 65, 65)
            SurfaceHover = [System.Drawing.Color]::FromArgb(60, 60, 60)
            Success = [System.Drawing.Color]::FromArgb(43, 194, 83)
            Warning = [System.Drawing.Color]::FromArgb(255, 191, 0)
            Error = [System.Drawing.Color]::FromArgb(232, 17, 35)
            MenuBar = [System.Drawing.Color]::FromArgb(35, 35, 35)
            StatusBar = [System.Drawing.Color]::FromArgb(0, 120, 215)
        }
    } else {
        return @{
            Background = [System.Drawing.Color]::White
            Surface = [System.Drawing.Color]::FromArgb(249, 249, 249)
            Text = [System.Drawing.Color]::Black
            TextSecondary = [System.Drawing.Color]::FromArgb(150, 0, 0, 0)
            Border = [System.Drawing.Color]::FromArgb(220, 220, 220)
            Primary = [System.Drawing.Color]::FromArgb(128, 0, 128)  # Purple color
            PrimaryHover = [System.Drawing.Color]::FromArgb(100, 0, 100)  # Darker purple
            Secondary = [System.Drawing.Color]::FromArgb(240, 240, 240)
            SurfaceHover = [System.Drawing.Color]::FromArgb(235, 235, 235)
            Success = [System.Drawing.Color]::FromArgb(43, 194, 83)
            Warning = [System.Drawing.Color]::FromArgb(255, 191, 0)
            Error = [System.Drawing.Color]::FromArgb(232, 17, 35)
            MenuBar = [System.Drawing.Color]::FromArgb(245, 245, 245)
            StatusBar = [System.Drawing.Color]::FromArgb(128, 0, 128)  # Purple color
        }
    }
}

# -------------------------
# Proxy Settings
# -------------------------
function Set-ProxySettings {
    param(
        [string]$ProxyAddr,
        [bool]$Enable
    )
    $global:ProxySettings.Enabled = $Enable
    $global:ProxySettings.Address = if ($Enable) { $ProxyAddr } else { "" }
    Export-Settings
}

# -------------------------
# Scheduled Tasks
# -------------------------
function Get-ScheduledUpdateTask {
    try {
        $task = Get-ScheduledTask -TaskName "RikorDriverUpdate" -ErrorAction SilentlyContinue
        return $task
    } catch {
        return $null
    }
}

function Set-ScheduledUpdate {
    param(
        [string]$Frequency = "Daily",
        [string]$Time = "03:00"
    )
    try {
        # Remove existing task if it exists
        Unregister-ScheduledTask -TaskName "RikorDriverUpdate" -Confirm:$false -ErrorAction SilentlyContinue
        
        # Parse time
        $timeParts = $Time.Split(':')
        $hour = [int]$timeParts[0]
        $minute = [int]$timeParts[1]
        
        # Create new scheduled task
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
        $trigger = switch ($Frequency) {
            "Daily" { New-ScheduledTaskTrigger -Daily -At "$Time" }
            "Weekly" { New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At "$Time" }
            "Monthly" { New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Monday -At "$Time" }
        }
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        Register-ScheduledTask -TaskName "RikorDriverUpdate" -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Description "Automated Rikor Driver Updates"
        return $true
    } catch {
        Write-Warning "Failed to set scheduled update: $_"
        return $false
    }
}

function Remove-ScheduledUpdate {
    try {
        Unregister-ScheduledTask -TaskName "RikorDriverUpdate" -Confirm:$false -ErrorAction SilentlyContinue
        return $true
    } catch {
        Write-Warning "Failed to remove scheduled update: $_"
        return $false
    }
}

# -------------------------
# System Restore Point
# -------------------------
function New-RestorePoint {
    param([string]$Description)
    try {
        $dt = Get-Date
        $params = @{
            Description = $Description
            RestorePointType = "MODIFY_SETTINGS"
            EventType = "BEGIN_SYSTEM_CHANGE"
        }
        Checkpoint-Computer @params
        return $true
    } catch {
        Write-Warning "Failed to create restore point: $_"
        return $false
    }
}

# -------------------------
# Main Form
# -------------------------
$form = New-Object Windows.Forms.Form
$form.Text = $AppTitle
$form.Size = '800,600'
$form.StartPosition = "CenterScreen"
$form.MinimizeBox = $true
$form.MaximizeBox = $true
$form.Icon = [System.Drawing.SystemIcons]::Shield
$form.Font = New-Object Drawing.Font("Segoe UI", 9.5)

# Menu Strip
$menuStrip = New-Object Windows.Forms.MenuStrip
$form.Controls.Add($menuStrip)

# File menu
$fileMenu = New-Object Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "&File"
$menuStrip.Items.Add($fileMenu)

$menuOpenLogs = New-Object Windows.Forms.ToolStripMenuItem
$menuOpenLogs.Text = (Get-LocalizedString "BtnOpenLogs")
$fileMenu.DropDownItems.Add($menuOpenLogs)

$menuExit = New-Object Windows.Forms.ToolStripMenuItem
$menuExit.Text = "E&xit"
$fileMenu.DropDownItems.Add($menuExit)

# Update menu
$updateMenu = New-Object Windows.Forms.ToolStripMenuItem
$updateMenu.Text = "&Update"
$menuStrip.Items.Add($updateMenu)

$menuWU = New-Object Windows.Forms.ToolStripMenuItem
$menuWU.Text = (Get-LocalizedString "BtnWU")
$updateMenu.DropDownItems.Add($menuWU)

$menuCheckUpdates = New-Object Windows.Forms.ToolStripMenuItem
$menuCheckUpdates.Text = (Get-LocalizedString "BtnCheckUpdates")
$updateMenu.DropDownItems.Add($menuCheckUpdates)

$menuRikorUpdate = New-Object Windows.Forms.ToolStripMenuItem  # NEW: Rikor update menu item
$menuRikorUpdate.Text = (Get-LocalizedString "BtnRikorUpdate")
$updateMenu.DropDownItems.Add($menuRikorUpdate)

$menuDownloadAndInstall = New-Object Windows.Forms.ToolStripMenuItem
$menuDownloadAndInstall.Text = (Get-LocalizedString "BtnDownloadAndInstall")
$updateMenu.DropDownItems.Add($menuDownloadAndInstall)

# Tools menu
$toolsMenu = New-Object Windows.Forms.ToolStripMenuItem
$toolsMenu.Text = "&Tools"
$menuStrip.Items.Add($toolsMenu)

$menuScan = New-Object Windows.Forms.ToolStripMenuItem
$menuScan.Text = (Get-LocalizedString "BtnScan")
$toolsMenu.DropDownItems.Add($menuScan)

$menuBackup = New-Object Windows.Forms.ToolStripMenuItem
$menuBackup.Text = (Get-LocalizedString "BtnBackup")
$toolsMenu.DropDownItems.Add($menuBackup)

$menuInstall = New-Object Windows.Forms.ToolStripMenuItem
$menuInstall.Text = (Get-LocalizedString "BtnInstall")
$toolsMenu.DropDownItems.Add($menuInstall)

$menuCancel = New-Object Windows.Forms.ToolStripMenuItem
$menuCancel.Text = (Get-LocalizedString "BtnCancel")
$toolsMenu.DropDownItems.Add($menuCancel)

# Options menu
$optionsMenu = New-Object Windows.Forms.ToolStripMenuItem
$optionsMenu.Text = "&Options"
$menuStrip.Items.Add($optionsMenu)

$menuRestorePoint = New-Object Windows.Forms.ToolStripMenuItem
$menuRestorePoint.Text = (Get-LocalizedString "BtnRestorePoint")
$optionsMenu.DropDownItems.Add($menuRestorePoint)

$menuSchedule = New-Object Windows.Forms.ToolStripMenuItem
$menuSchedule.Text = (Get-LocalizedString "BtnSchedule")
$optionsMenu.DropDownItems.Add($menuSchedule)

$menuFilters = New-Object Windows.Forms.ToolStripMenuItem
$menuFilters.Text = (Get-LocalizedString "BtnFilters")
$optionsMenu.DropDownItems.Add($menuFilters)

$menuHistory = New-Object Windows.Forms.ToolStripMenuItem
$menuHistory.Text = (Get-LocalizedString "BtnHistory")
$optionsMenu.DropDownItems.Add($menuHistory)

$menuSettingsTop = New-Object Windows.Forms.ToolStripMenuItem
$menuSettingsTop.Text = (Get-LocalizedString "BtnSettings")
$optionsMenu.DropDownItems.Add($menuSettingsTop)

$menuToggleTheme = New-Object Windows.Forms.ToolStripMenuItem
$menuToggleTheme.Text = if ($global:DarkModeEnabled) { (Get-LocalizedString "BtnLightMode") } else { (Get-LocalizedString "BtnDarkMode") }
$optionsMenu.DropDownItems.Add($menuToggleTheme)

# Language menu
$menuLanguage = New-Object Windows.Forms.ToolStripMenuItem
$menuLanguage.Text = "&Language"
$optionsMenu.DropDownItems.Add($menuLanguage)

# Add only English and Russian to language menu
$langEn = New-Object Windows.Forms.ToolStripMenuItem
$langEn.Text = "English"
$langEn.Tag = "en"
$langEn.Checked = ($global:CurrentLanguage -eq "en")
$menuLanguage.DropDownItems.Add($langEn)

$langRu = New-Object Windows.Forms.ToolStripMenuItem
$langRu.Text = "Русский"
$langRu.Tag = "ru"
$langRu.Checked = ($global:CurrentLanguage -eq "ru")
$menuLanguage.DropDownItems.Add($langRu)

# Toolbar Panel
$toolbarPanel = New-Object Windows.Forms.Panel
$toolbarPanel.Dock = 'Top'
$toolbarPanel.Height = 100
$toolbarPanel.Padding = '10,10,10,10'
$form.Controls.Add($toolbarPanel)

# Button Container (for centering)
$buttonContainer = New-Object Windows.Forms.Panel
$buttonContainer.Dock = 'Fill'
$buttonContainer.BackColor = [System.Drawing.Color]::Transparent
$toolbarPanel.Controls.Add($buttonContainer)

# Toolbar Separator
$toolbarSeparator = New-Object Windows.Forms.Panel
$toolbarSeparator.Dock = 'Bottom'
$toolbarSeparator.Height = 1
$toolbarPanel.Controls.Add($toolbarSeparator)

# Toolbar buttons
# NEW: Adding Rikor Update button
$btnRikorUpdate = New-Object Windows.Forms.Button
$btnRikorUpdate.Text = Get-LocalizedString "BtnRikorUpdate"
$btnRikorUpdate.Size = '150,35'
$btnRikorUpdate.Margin = '5,5,5,5'
$btnRikorUpdate.FlatStyle = 'Flat'
$btnRikorUpdate.FlatAppearance.BorderSize = 0
$btnRikorUpdate.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnRikorUpdate.Tag = $true  # Mark as primary button
$buttonContainer.Controls.Add($btnRikorUpdate)

$btnWU = New-Object Windows.Forms.Button
$btnWU.Text = Get-LocalizedString "BtnWU"
$btnWU.Size = '150,35'
$btnWU.Margin = '5,5,5,5'
$btnWU.FlatStyle = 'Flat'
$btnWU.FlatAppearance.BorderSize = 0
$btnWU.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnWU.Tag = $false  # Mark as secondary button
$buttonContainer.Controls.Add($btnWU)

$btnCheckUpdates = New-Object Windows.Forms.Button
$btnCheckUpdates.Text = Get-LocalizedString "BtnCheckUpdates"
$btnCheckUpdates.Size = '150,35'
$btnCheckUpdates.Margin = '5,5,5,5'
$btnCheckUpdates.FlatStyle = 'Flat'
$btnCheckUpdates.FlatAppearance.BorderSize = 0
$btnCheckUpdates.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCheckUpdates.Tag = $false
$buttonContainer.Controls.Add($btnCheckUpdates)

$btnDownloadAndInstall = New-Object Windows.Forms.Button
$btnDownloadAndInstall.Text = Get-LocalizedString "BtnDownloadAndInstall"
$btnDownloadAndInstall.Size = '150,35'
$btnDownloadAndInstall.Margin = '5,5,5,5'
$btnDownloadAndInstall.FlatStyle = 'Flat'
$btnDownloadAndInstall.FlatAppearance.BorderSize = 0
$btnDownloadAndInstall.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnDownloadAndInstall.Tag = $false
$buttonContainer.Controls.Add($btnDownloadAndInstall)

$btnCancel = New-Object Windows.Forms.Button
$btnCancel.Text = Get-LocalizedString "BtnCancel"
$btnCancel.Size = '150,35'
$btnCancel.Margin = '5,5,5,5'
$btnCancel.FlatStyle = 'Flat'
$btnCancel.FlatAppearance.BorderSize = 0
$btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCancel.Tag = $false
$buttonContainer.Controls.Add($btnCancel)

# Center buttons function
function Update-ButtonContainerPadding {
    if ($buttonContainer.Width -gt 0) {
        $totalWidth = 0
        foreach ($btn in $buttonContainer.Controls) {
            if ($btn -is [System.Windows.Forms.Button]) {
                $totalWidth += $btn.Width
            }
        }
        $spacing = [Math]::Max(0, ($buttonContainer.Width - $totalWidth) / ($buttonContainer.Controls.Count + 1))
        $buttonContainer.Padding = [System.Windows.Forms.Padding]::new([int]$spacing, 10, [int]$spacing, 10)
    }
}

# Content Panel
$contentPanel = New-Object Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$form.Controls.Add($contentPanel)

# Header Label
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

# Status textbox with border
$statusBorderPanel = New-Object Windows.Forms.Panel
$statusBorderPanel.Dock = 'Fill'
$contentPanel.Controls.Add($statusBorderPanel)

$status = New-Object Windows.Forms.RichTextBox
$status.Multiline = $true
$status.ReadOnly = $true
$status.ScrollBars = 'Vertical'
$status.Font = New-Object Drawing.Font("Consolas", 8)
$status.Dock = 'Fill'
$statusBorderPanel.Controls.Add($status)

# Progress panel with border
$progressPanel = New-Object Windows.Forms.Panel
$progressPanel.Dock = 'Bottom'
$progressPanel.Height = 30
$progressPanel.Padding = '5,5,5,5'
$form.Controls.Add($progressPanel)

$progressBorderPanel = New-Object Windows.Forms.Panel
$progressBorderPanel.Dock = 'Fill'
$progressPanel.Controls.Add($progressBorderPanel)

$progress = New-Object Windows.Forms.ProgressBar
$progress.Dock = 'Fill'
$progressBorderPanel.Controls.Add($progress)

# Status bar
$statusBar = New-Object Windows.Forms.StatusStrip
$form.Controls.Add($statusBar)

$statusLabel = New-Object Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusLabel.Spring = $true
$statusBar.Items.Add($statusLabel)

$versionLabel = New-Object Windows.Forms.ToolStripStatusLabel
$versionLabel.Text = "v2.0"
$statusBar.Items.Add($versionLabel)

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
        
        try {
            # Execute task based on task name
            switch ($taskName) {
                "WindowsUpdate" {
                    L "Checking for Windows updates..."
                    try {
                        # Create update session
                        $session = New-Object -ComObject Microsoft.Update.Session
                        $searcher = $session.CreateUpdateSearcher()
                        
                        # Search for updates
                        L "Searching for applicable updates..."
                        $results = $searcher.Search("IsInstalled=0 and Type='Software'")
                        
                        if ($results.Updates.Count -gt 0) {
                            L "$(Get-LocalizedString 'UpdatesAvailable') $($results.Updates.Count) updates found"
                            
                            # Create downloader
                            $downloader = $session.CreateUpdateDownloader()
                            $downloader.Updates = $results.Updates
                            
                            # Download updates
                            L "Downloading updates..."
                            $downloader.Download()
                            
                            # Create installer
                            $installer = $session.CreateUpdateInstaller()
                            $installer.Updates = $results.Updates
                            
                            # Install updates
                            L "Installing updates..."
                            $installationResult = $installer.Install()
                            
                            # Corrected: Using if-else instead of ternary operator
                            if ($installationResult.ResultCode -eq 2) {
                                L "$(Get-LocalizedString 'UpdateSuccess')"
                                $rebootStatus = if ($installationResult.RebootRequired) { "Reboot required" } else { "No reboot required" }
                                L "Status: Updates installed - $rebootStatus"
                            } else {
                                L "[ERROR] $(Get-LocalizedString 'UpdateFailed') - Result code: $($installationResult.ResultCode)"
                            }
                        } else {
                            L "$(Get-LocalizedString 'NoUpdatesAvailable')"
                        }
                    } catch {
                        L "[ERROR] Failed to check/install Windows updates: $_"
                    }
                    L "Completed"
                }
                
                "MicrosoftUpdate" {
                    L "Checking for Microsoft updates..."
                    try {
                        # This would typically check for Microsoft product updates beyond Windows
                        # Implementation depends on specific Microsoft update sources
                        L "Microsoft update check completed (stub implementation)"
                    } catch {
                        L "[ERROR] Failed to check Microsoft updates: $_"
                    }
                    L "Completed"
                }
                
                "RikorUpdate" {
                    L "Checking for Rikor updates..."
                    try {
                        # Check if Rikor update server is available
                        if (Test-RikorUpdates) {
                            L "Rikor update server is available"
                            # Actual implementation would go here
                            L "Rikor update check completed (actual implementation needed)"
                        } else {
                            L "$(Get-LocalizedString 'RikorServerUnavailable')"
                        }
                    } catch {
                        L "[ERROR] Failed to check Rikor updates: $_"
                    }
                    L "Completed"
                }
                
                "DownloadAndInstallDrivers" {
                    # Define the public ZIP URL here (REPLACE WITH ACTUAL LINK YOU GET FROM NEXTCLOUD SHARE OR GOOGLE DRIVE)
                    # Example for Google Drive: $zipUrl = "https://drive.google.com/uc?export=download&id=FILE_ID"
                    $zipUrl = "https://drive.google.com/uc?export=download&id=14_iaT8zdS800GpL76CSVb5vBQN7whZ8w" # <--- INSERTED YOUR GOOGLE DRIVE LINK

                    # Validate the URL
                    if (-not $zipUrl -or $zipUrl -eq "https://drive.google.com/uc?export=download&id=14_iaT8zdS800GpL76CSVb5vBQN7whZ8w") {
                        L "[ERROR] Public ZIP download URL is not configured. Please edit the script."
                        L "Replace 'https://drive.google.com/uc?export=download&id=14_iaT8zdS800GpL76CSVb5vBQN7whZ8w' with the actual link."
                        L "Completed"
                        return
                    }

                    # Create temp directory
                    $tempDir = Join-Path $env:TEMP "RikorDriversTemp_$(Get-Date -Format 'yyyyMMddHHmmss')"
                    try {
                        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                        L "Created temporary directory: $tempDir"

                        # Define zip file path
                        $zipPath = Join-Path $tempDir "drivers.zip"

                        # Download ZIP
                        L "$(Get-LocalizedString 'CheckingForUpdates') drivers archive from: $zipUrl"
                        try {
                            # Use basic parsing to avoid issues with complex pages
                            Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
                            L "Download completed to: $zipPath"
                        } catch {
                            L "[ERROR] Failed to download ZIP: $_"
                            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                            L "Completed"
                            return
                        }

                        # Check if ZIP exists and is not empty
                        if (-not (Test-Path $zipPath) -or (Get-Item $zipPath).Length -eq 0) {
                            L "[ERROR] Downloaded ZIP file is missing or empty."
                            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                            L "Completed"
                            return
                        }

                        # Extract ZIP
                        $extractDir = Join-Path $tempDir "ExtractedDrivers"
                        L "Extracting ZIP archive to: $extractDir"
                        try {
                            Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
                            L "Extraction completed."
                        } catch {
                            L "[ERROR] Failed to extract ZIP: $_"
                            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                            L "Completed"
                            return
                        }

                        # Verify extracted content
                        if (-not (Test-Path $extractDir)) {
                             L "[ERROR] Extraction directory '$extractDir' does not exist after extraction."
                             Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                             L "Completed"
                             return
                        }

                        $infFilesInZip = Get-ChildItem -Path $extractDir -Recurse -Include "*.inf" -ErrorAction SilentlyContinue
                        if ($infFilesInZip.Count -eq 0) {
                            L "[WARNING] No .inf files found in the downloaded archive after extraction."
                            # Still proceed to InstallDrivers task, which will report its own error if needed
                        } else {
                            L "Found $($infFilesInZip.Count) .inf file(s) in the archive."
                        }

                        # Prepare arguments for the InstallDrivers task logic
                        $installArgs = @($extractDir)

                        # --- Simulate calling the InstallDrivers logic directly within the same job ---
                        # This avoids creating another nested job and keeps logging consistent.
                        $folder = $installArgs[0]
                        L "$(Get-LocalizedString 'InstallingUpdates') drivers from extracted archive: $folder"
                        try {
                            if (-not (Test-Path $folder)) {
                                L "[ERROR] Folder not found: $folder (This should not happen after extraction)"
                                L "Completed"
                                return
                            }

                            # Find INF files recursively in the extracted folder
                            $infFiles = Get-ChildItem -Path $folder -Recurse -Include *.inf -ErrorAction SilentlyContinue

                            if ($infFiles.Count -eq 0) {
                                L "[ERROR] No .inf driver files found in the extracted folder: $folder"
                                L "Completed"
                                return
                            }

                            L "Found $($infFiles.Count) driver file(s) in extracted folder"
                            L "Installing drivers..."
                            L ""
                            $successCount = 0
                            $failCount = 0
                            $current = 0
                            foreach ($inf in $infFiles) {
                                $current++
                                L "[$current/$($infFiles.Count)] $($inf.Name)"
                                try {
                                    # Use pnputil to add and install the driver
                                    $out = & pnputil.exe /add-driver $inf.FullName /install /force 2>&1
                                    $hasError = $false
                                    $out | ForEach-Object {
                                        # Check for common failure indicators in pnputil output
                                        if ($_ -match "(error|failed|fail|cannot find suitable)" -and -not $hasError) {
                                            L "     $_" # Log the error line
                                            $hasError = $true
                                            $failCount++
                                        } elseif ($_ -match "^Published the driver") {
                                             # Successfully added to driver store
                                             # Actual installation success is harder to determine without checking exit code,
                                             # but if no explicit error occurred, assume success for counting.
                                             # pnputil exit code 0 usually means added/installed ok, non-zero means error.
                                        }
                                    }
                                    if (-not $hasError) {
                                         # If no specific error was logged, check the exit code
                                         if ($LASTEXITCODE -eq 0) {
                                              $successCount++
                                              L "     -> Added and installed successfully." # Optional verbose log
                                         } else {
                                              # pnputil reported an error via exit code, even if not captured in output
                                              $failCount++
                                              L "     -> Installation failed (pnputil exit code: $LASTEXITCODE)." # Log exit code
                                         }
                                    }
                                } catch {
                                    $failCount++
                                    L "     -> Exception during installation: $_"
                                }
                                Start-Sleep -Milliseconds 300
                            }
                            L ""
                            L "Installation from archive complete: $successCount successful, $failCount failed"
                            # Corrected: Using if-else instead of ternary operator
                            if ($successCount -gt 0) {
                                $rebootRequired = if (Get-Variable -Name installResult -ErrorAction SilentlyContinue) {
                                    if ($installResult.RebootRequired) { "Reboot required" } else { "No reboot required" }
                                } else {
                                    "Note: Reboot may be required for some drivers to take effect."
                                }
                                L $rebootRequired
                            }
                        } catch {
                            L "[ERROR] Installation process failed: $_"
                        }
                        # --- End of simulated InstallDrivers logic ---

                    } catch {
                        L "[ERROR] An unexpected error occurred during download/extraction: $_"
                    } finally {
                        # Always attempt to clean up the temporary directory
                        L "Cleaning up temporary directory: $tempDir"
                        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                        if (Test-Path $tempDir) {
                            L "[WARNING] Could not remove temporary directory: $tempDir"
                        } else {
                            L "Temporary directory cleaned up successfully."
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
                        foreach ($inf in $infFiles) {
                            $current++
                            L "[$current/$($infFiles.Count)] $($inf.Name)"
                            try {
                                $out = & pnputil.exe /add-driver $inf.FullName /install /force 2>&1
                                $hasError = $false
                                $out | ForEach-Object {
                                    if ($_ -match "(error|failed|fail|cannot find suitable)" -and -not $hasError) {
                                        L "     $_" # Log the error line
                                        $hasError = $true
                                        $failCount++
                                    }
                                }
                                if (-not $hasError) {
                                    # If no specific error was logged, check the exit code
                                    if ($LASTEXITCODE -eq 0) {
                                         $successCount++
                                         L "     -> Added and installed successfully." # Optional verbose log
                                    } else {
                                         # pnputil reported an error via exit code, even if not captured in output
                                         $failCount++
                                         L "     -> Installation failed (pnputil exit code: $LASTEXITCODE)." # Log exit code
                                    }
                                }
                            } catch {
                                $failCount++
                                L "     -> Exception during installation: $_"
                            }
                            Start-Sleep -Milliseconds 300
                        }
                        L ""
                        L "Installation complete: $successCount successful, $failCount failed"
                        # Corrected: Using if-else instead of ternary operator
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
            if ($contentStr -match "$(Get-LocalizedString 'CheckingForUpdates')") { $p = 5 }
            if ($contentStr -match "Download completed") { $p = 15 }
            if ($contentStr -match "Extracting ZIP archive") { $p = 20 }
            if ($contentStr -match "Extraction completed") { $p = 30 }
            if ($contentStr -match "$(Get-LocalizedString 'InstallingUpdates')") { $p = 40 }
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
            if ($contentStr -match "$(Get-LocalizedString 'UpdateSuccess')") { $p = 95 }
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
    $btnWU.Text = Get-LocalizedString "BtnWU"
    $btnCheckUpdates.Text = Get-LocalizedString "BtnCheckUpdates"
    $btnRikorUpdate.Text = Get-LocalizedString "BtnRikorUpdate"  # NEW: Update Rikor button text
    $btnDownloadAndInstall.Text = Get-LocalizedString "BtnDownloadAndInstall"
    $btnScan.Text = Get-LocalizedString "BtnScan"
    $btnBackup.Text = Get-LocalizedString "BtnBackup"
    $btnInstall.Text = Get-LocalizedString "BtnInstall"
    $btnCancel.Text = Get-LocalizedString "BtnCancel"
    # Menu items
    $menuOpenLogs.Text = Get-LocalizedString "BtnOpenLogs"
    $menuWU.Text = Get-LocalizedString "BtnWU"
    $menuCheckUpdates.Text = Get-LocalizedString "BtnCheckUpdates"
    $menuRikorUpdate.Text = Get-LocalizedString "BtnRikorUpdate"  # NEW: Update Rikor menu item text
    $menuDownloadAndInstall.Text = Get-LocalizedString "BtnDownloadAndInstall"
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
function Invoke-WindowsUpdate {
    $status.Clear()
    $progress.Value = 0
    $statusLabel.Text = "  Checking for Windows updates..."
    Start-BackgroundTask -Name "WindowsUpdate" -TaskArgs @()
}

function Invoke-MicrosoftUpdate {
    $status.Clear()
    $progress.Value = 0
    $statusLabel.Text = "  Checking for Microsoft updates..."
    Start-BackgroundTask -Name "MicrosoftUpdate" -TaskArgs @()
}

function Invoke-RikorUpdate {
    $status.Clear()
    $progress.Value = 0
    $statusLabel.Text = "  Checking for Rikor updates..."
    Start-BackgroundTask -Name "RikorUpdate" -TaskArgs @()
}

function Invoke-DownloadAndInstallDrivers {
    $status.Clear()
    $progress.Value = 0
    $statusLabel.Text = "  Downloading and installing drivers from Rikor..."
    Start-BackgroundTask -Name "DownloadAndInstallDrivers" -TaskArgs @()
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
$btnWU.Add_Click({ Invoke-WindowsUpdate })
$btnCheckUpdates.Add_Click({ Invoke-MicrosoftUpdate })
$btnRikorUpdate.Add_Click({ Invoke-RikorUpdate })  # NEW: Rikor update button handler
$btnDownloadAndInstall.Add_Click({ Invoke-DownloadAndInstallDrivers })
$btnScan.Add_Click({ Invoke-ScanDrivers })
$btnBackup.Add_Click({ Invoke-BackupDrivers })
$btnInstall.Add_Click({ Invoke-InstallDrivers })
$btnCancel.Add_Click({ Invoke-CancelTask })

# -------------------------
# Menu handlers
# -------------------------
$menuWU.Add_Click({ Invoke-WindowsUpdate })
$menuCheckUpdates.Add_Click({ Invoke-MicrosoftUpdate })
$menuRikorUpdate.Add_Click({ Invoke-RikorUpdate })  # NEW: Rikor update menu handler
$menuDownloadAndInstall.Add_Click({ Invoke-DownloadAndInstallDrivers })
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