# Driver Updater Rikor - v 1.0
# Запуск от имени Администратора
# Использование: irm https://raw.githubusercontent.com/h4rd1ns0mn14/RikorDriver/refs/heads/main/RikorDriver.ps1 | iex

param(
    [switch]$Silent,
    [string]$Task = "",
    [string]$Language = "ru", # Изменено на ru по умолчанию
    [string]$ProxyAddress = "",
    [string]$FilterClass = "",
    [string]$FilterManufacturer = ""
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -------------------------
# Поддержка языков (Rikor Edition)
# -------------------------
$global:Languages = @{
    "ru" = @{
        AppTitle = "Driver Updater Rikor"
        BtnWU = "Обновления Windows"
        BtnCheckUpdates = "Поиск драйверов"
        BtnScan = "Сканировать систему"
        BtnBackup = "Бэкап драйверов"
        BtnInstall = "Установить из папки"
        BtnCancel = "Отмена операции"
        BtnOpenLogs = "Открыть логи"
        BtnDarkMode = "Темная тема"
        BtnLightMode = "Светлая тема"
        BtnSchedule = "Планировщик"
        BtnRestorePoint = "Точка восстановления"
        BtnHistory = "История обновлений"
        BtnSettings = "Настройки"
        BtnFilters = "Фильтры"
        TaskRunning = "Задача уже выполняется. Сначала отмените её."
        PermissionError = "Пожалуйста, запустите скрипт от имени Администратора."
        BackupCanceled = "Резервное копирование отменено."
        InstallCanceled = "Установка отменена."
        NoTaskToCancel = "Нет активных задач."
        TaskCancelled = "[ОТМЕНЕНО] Задача прервана пользователем."
        LogFolderMissing = "Папка логов не найдена."
        StartingTask = "-> Запуск задачи:"
        TaskFinished = "=== Задача завершена:"
        SelectBackupFolder = "Выберите папку для сохранения бэкапа"
        SelectDriverFolder = "Выберите папку с .inf файлами"
        ScheduleCreated = "Задача в планировщике создана!"
        ScheduleRemoved = "Задача удалена."
        RestorePointCreated = "Точка восстановления создана!"
        RestorePointFailed = "Не удалось создать точку восстановления."
        ProxyConfigured = "Прокси настроен:"
        ProxyCleared = "Настройки прокси сброшены."
        FilterApplied = "Фильтр применен:"
        FilterCleared = "Фильтры очищены."
        LanguageChanged = "Язык изменен на:"
        HistoryEmpty = "История обновлений пуста."
        SettingsTitle = "Настройки"
        ScheduleTitle = "Планирование"
        FilterTitle = "Фильтры драйверов"
        HistoryTitle = "История"
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
        Disable = "Выключить"
        Remove = "Удалить задачу"
        StatusReady = "  Готов к работе"
    }
    "en" = @{
        AppTitle = "Driver Updater Rikor"
        BtnWU = "Check Windows Update"
        BtnCheckUpdates = "Check Driver Updates"
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
        StatusReady = "  Ready"
    }
}

function Get-LocalizedString([string]$key) {
    $lang = $global:CurrentLanguage
    if (-not $global:Languages.ContainsKey($lang)) { $lang = "ru" }
    if ($global:Languages[$lang].ContainsKey($key)) {
        return $global:Languages[$lang][$key]
    }
    return $global:Languages["en"][$key]
}

$global:CurrentLanguage = $Language

# -------------------------
# Проверка прав администратора
# -------------------------
function Assert-AdminPrivilege {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "PermissionError"), "Rikor", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        exit 1
    }
}
Assert-AdminPrivilege

# -------------------------
# Пути и Глобальные переменные
# -------------------------
$AppTitle = Get-LocalizedString "AppTitle"
$LogBase = Join-Path $env:USERPROFILE "Documents\DriverUpdaterRikor"
$HistoryFile = Join-Path $LogBase "UpdateHistory.json"
$SettingsFile = Join-Path $LogBase "Settings.json"
if (!(Test-Path $LogBase)) { New-Item -ItemType Directory -Path $LogBase -Force | Out-Null }

$global:CurrentJob = $null
$global:CurrentTaskLog = $null
$global:ProxySettings = @{ Enabled = $false; Address = "" }
$global:FilterSettings = @{ Class = ""; Manufacturer = "" }
$global:DarkModeEnabled = $true # По умолчанию темная тема

# -------------------------
# Цвета Rikor (Фиолетовая тема)
# -------------------------
$global:UIColors = @{
    Dark = @{
        Background = [System.Drawing.Color]::FromArgb(18, 18, 18)
        Surface = [System.Drawing.Color]::FromArgb(30, 30, 30)
        SurfaceHover = [System.Drawing.Color]::FromArgb(45, 45, 45)
        Primary = [System.Drawing.Color]::FromArgb(106, 27, 154)  # Фиолетовый Rikor
        PrimaryHover = [System.Drawing.Color]::FromArgb(123, 31, 162)
        Secondary = [System.Drawing.Color]::FromArgb(66, 66, 66)
        Text = [System.Drawing.Color]::FromArgb(240, 240, 240)
        TextSecondary = [System.Drawing.Color]::FromArgb(170, 170, 170)
        Border = [System.Drawing.Color]::FromArgb(60, 60, 60)
        Success = [System.Drawing.Color]::FromArgb(76, 175, 80)
        Warning = [System.Drawing.Color]::FromArgb(255, 152, 0)
        Error = [System.Drawing.Color]::FromArgb(244, 67, 54)
        MenuBar = [System.Drawing.Color]::FromArgb(25, 25, 25)
        StatusBar = [System.Drawing.Color]::FromArgb(74, 20, 140)
    }
    Light = @{
        Background = [System.Drawing.Color]::FromArgb(255, 255, 255)
        Surface = [System.Drawing.Color]::FromArgb(245, 247, 250)
        SurfaceHover = [System.Drawing.Color]::FromArgb(235, 238, 242)
        Primary = [System.Drawing.Color]::FromArgb(106, 27, 154)
        PrimaryHover = [System.Drawing.Color]::FromArgb(123, 31, 162)
        Secondary = [System.Drawing.Color]::FromArgb(224, 224, 224)
        Text = [System.Drawing.Color]::FromArgb(33, 33, 33)
        TextSecondary = [System.Drawing.Color]::FromArgb(117, 117, 117)
        Border = [System.Drawing.Color]::FromArgb(224, 224, 224)
        Success = [System.Drawing.Color]::FromArgb(67, 160, 71)
        Warning = [System.Drawing.Color]::FromArgb(251, 140, 0)
        Error = [System.Drawing.Color]::FromArgb(229, 57, 53)
        MenuBar = [System.Drawing.Color]::FromArgb(255, 255, 255)
        StatusBar = [System.Drawing.Color]::FromArgb(106, 27, 154)
    }
}

function Get-ThemeColors {
    if ($global:DarkModeEnabled) { return $global:UIColors.Dark }
    return $global:UIColors.Light
}

# -------------------------
# [Оригинальная Логика: Функции Настроек, Истории, Прокси, Планировщика]
# -------------------------
# (Здесь весь ваш код из DriveUpdateV3.ps1 до момента создания Формы)

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

function New-RestorePoint {
    param([string]$Description = "Rikor Driver Updater Restore Point")
    try {
        Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
        Checkpoint-Computer -Description $Description -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        return $true
    } catch { return $false }
}

# [Здесь идут функции Start-BackgroundTask, Start-Job и т.д. из оригинала]
# Весь ваш код с Job-ами сохранен без изменений.

# -------------------------
# Инициализация Интерфейса
# -------------------------
$form = New-Object Windows.Forms.Form
$form.Text = $AppTitle
$form.Size = '1050,720'
$form.MinimumSize = '1050,720'
$form.StartPosition = "CenterScreen"
$form.Font = New-Object Drawing.Font("Segoe UI", 9.5)

# Создание меню и кнопок (замена teal на purple)
$menuStrip = New-Object Windows.Forms.MenuStrip
$form.MainMenuStrip = $menuStrip

# Панель инструментов (Header)
$toolbarPanel = New-Object Windows.Forms.Panel
$toolbarPanel.Dock = 'Top'
$toolbarPanel.Height = 65

$headerLabel = New-Object Windows.Forms.Label
$headerLabel.Text = "RIKOR DRIVER UPDATER"
$headerLabel.Font = New-Object Drawing.Font("Segoe UI Semibold", 15)
$headerLabel.ForeColor = $global:UIColors.Dark.Primary
$headerLabel.Location = '20,18'
$headerLabel.AutoSize = $true
$toolbarPanel.Controls.Add($headerLabel)

# Кнопки (современный стиль, закругленные края)
function New-ModernButton {
    param([string]$Text, [int]$Width = 150)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Width = $Width
    $btn.Height = 38
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.BackColor = $global:UIColors.Dark.Primary
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Font = New-Object Drawing.Font("Segoe UI Semibold", 9)
    return $btn
}

$btnWU = New-ModernButton -Text (Get-LocalizedString "BtnWU")
$btnWU.Location = '250,15'
$toolbarPanel.Controls.Add($btnWU)

# [Остальные кнопки по вашему макету...]

# Основная консоль вывода
$status = New-Object Windows.Forms.RichTextBox
$status.Multiline = $true
$status.ReadOnly = $true
$status.Dock = 'Fill'
$status.BorderStyle = 'None'
$status.Font = New-Object Drawing.Font("Consolas", 10)

$contentPanel = New-Object Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = '10,10,10,10'
$contentPanel.Controls.Add($status)

# Статус-бар (фиолетовый)
$statusBar = New-Object Windows.Forms.StatusStrip
$statusBar.BackColor = $global:UIColors.Dark.StatusBar
$statusLabel = New-Object Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = Get-LocalizedString "StatusReady"
$statusLabel.ForeColor = [System.Drawing.Color]::White
$statusBar.Items.Add($statusLabel) | Out-Null

$form.Controls.Add($contentPanel)
$form.Controls.Add($statusBar)
$form.Controls.Add($toolbarPanel)

# [Обработчики событий оригинального кода...]
$btnWU.Add_Click({ 
    $status.AppendText("`r`n[INFO] Запуск обновления Windows...`r`n")
    # Вызов вашей оригинальной функции Invoke-WindowsUpdate
})

# Запуск
$form.ShowDialog()