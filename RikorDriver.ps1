# Driver Updater Rikor
# Run as Administrator

param(
    [switch]$Silent,
    [string]$Task = "",
    [string]$Language = "ru", # Изменено на русский по умолчанию
    [string]$ProxyAddress = "",
    [string]$FilterClass = "",
    [string]$FilterManufacturer = ""
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

# -------------------------
# Поддержка языков
# -------------------------
$global:Languages = @{
    "ru" = @{
        AppTitle = "Driver Updater Rikor"
        BtnWU = "Центр обновления Windows"
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
        StatusReady = " Готово"
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
        StatusReady = " Ready"
    }
}

function Get-LocalizedString([string]$key) {
    $lang = $global:CurrentLanguage
    if (-not $global:Languages.ContainsKey($lang)) { $lang = "ru" }
    if ($global:Languages[$lang].ContainsKey($key)) { return $global:Languages[$lang][$key] }
    return $global:Languages["en"][$key]
}

$global:CurrentLanguage = $Language

# -------------------------
# Проверка прав администратора
# -------------------------
function Assert-AdminPrivilege {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "PermissionError"), "Rikor Driver Updater", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }
}
Assert-AdminPrivilege

# -------------------------
# Глобальные переменные и пути
# -------------------------
$AppTitle = Get-LocalizedString "AppTitle"
$LogBase = Join-Path $env:USERPROFILE "Documents\Rikor_DriverUpdater"
if (!(Test-Path $LogBase)) { New-Item -ItemType Directory -Path $LogBase -Force | Out-Null }
$SettingsFile = Join-Path $LogBase "Settings.json"

$global:DarkModeEnabled = $true # По умолчанию темная тема
$global:ProxySettings = @{ Enabled = $false; Address = "" }
$global:FilterSettings = @{ Class = ""; Manufacturer = "" }

# -------------------------
# Цветовая схема (ФИОЛЕТОВАЯ)
# -------------------------
$global:UIColors = @{
    Dark = @{
        Background = [System.Drawing.Color]::FromArgb(18, 18, 18)
        Surface = [System.Drawing.Color]::FromArgb(30, 30, 30)
        Primary = [System.Drawing.Color]::FromArgb(106, 27, 154)  # Насыщенный фиолетовый
        PrimaryHover = [System.Drawing.Color]::FromArgb(123, 31, 162)
        Secondary = [System.Drawing.Color]::FromArgb(66, 66, 66)
        Text = [System.Drawing.Color]::FromArgb(240, 240, 240)
        StatusBar = [System.Drawing.Color]::FromArgb(74, 20, 140) # Темно-фиолетовый
    }
    Light = @{
        Background = [System.Drawing.Color]::FromArgb(255, 255, 255)
        Surface = [System.Drawing.Color]::FromArgb(245, 247, 250)
        Primary = [System.Drawing.Color]::FromArgb(106, 27, 154)
        PrimaryHover = [System.Drawing.Color]::FromArgb(123, 31, 162)
        Secondary = [System.Drawing.Color]::FromArgb(224, 224, 224)
        Text = [System.Drawing.Color]::FromArgb(33, 33, 33)
        StatusBar = [System.Drawing.Color]::FromArgb(106, 27, 154)
    }
}

function Get-ThemeColors {
    if ($global:DarkModeEnabled) { return $global:UIColors.Dark }
    return $global:UIColors.Light
}

# -------------------------
# Построение интерфейса
# -------------------------
$form = New-Object Windows.Forms.Form
$form.Text = $AppTitle
$form.Size = '1050,750'
$form.StartPosition = "CenterScreen"
$form.BackColor = (Get-ThemeColors).Background
$form.Font = New-Object Drawing.Font("Segoe UI", 9.5)

# Панель логотипа
$logoPanel = New-Object Windows.Forms.Panel
$logoPanel.Dock = 'Top'
$logoPanel.Height = 80
$logoPanel.BackColor = (Get-ThemeColors).Surface

$pictureBox = New-Object Windows.Forms.PictureBox
$pictureBox.Size = '200,60'
$pictureBox.Location = '20,10'
$pictureBox.SizeMode = 'Zoom'

# Загрузка логотипа
$logoUrl = "https://rikor.com/wp-content/uploads/2025/09/group.svg"
$tempLogo = Join-Path $env:TEMP "rikor_logo.svg"
try {
    Invoke-WebRequest -Uri $logoUrl -OutFile $tempLogo -ErrorAction SilentlyContinue
    # WinForms не читает SVG нативно, но мы оставляем контейнер для совместимости или замены на PNG
    # Если бы это был PNG, код ниже бы сработал:
    # $pictureBox.Image = [System.Drawing.Image]::FromFile($tempLogo)
} catch {}

$titleLabel = New-Object Windows.Forms.Label
$titleLabel.Text = "DRIVER UPDATER"
$titleLabel.Font = New-Object Drawing.Font("Segoe UI Semibold", 18)
$titleLabel.ForeColor = (Get-ThemeColors).Primary
$titleLabel.Location = '230,20'
$titleLabel.AutoSize = $true

$logoPanel.Controls.Add($pictureBox)
$logoPanel.Controls.Add($titleLabel)
$form.Controls.Add($logoPanel)

# Тулбар с кнопками
$toolbar = New-Object Windows.Forms.FlowLayoutPanel
$toolbar.Dock = 'Top'
$toolbar.Height = 60
$toolbar.Padding = '20,10,0,0'
$toolbar.BackColor = (Get-ThemeColors).Background

function New-RikorButton($Text, $Width = 160) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Width = $Width
    $btn.Height = 40
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.BackColor = (Get-ThemeColors).Primary
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Font = New-Object Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    return $btn
}

$btnWU = New-RikorButton (Get-LocalizedString "BtnWU") 200
$btnCheck = New-RikorButton (Get-LocalizedString "BtnCheckUpdates") 180
$btnScan = New-RikorButton (Get-LocalizedString "BtnScan") 180

$toolbar.Controls.AddRange(@($btnWU, $btnCheck, $btnScan))
$form.Controls.Add($toolbar)

# Консоль вывода
$statusBox = New-Object Windows.Forms.RichTextBox
$statusBox.Dock = 'Fill'
$statusBox.ReadOnly = $true
$statusBox.BackColor = (Get-ThemeColors).Surface
$statusBox.ForeColor = (Get-ThemeColors).Text
$statusBox.BorderStyle = 'None'
$statusBox.Font = New-Object Drawing.Font("Consolas", 10)
$form.Controls.Add($statusBox)

# Статус-бар
$statusBar = New-Object Windows.Forms.Panel
$statusBar.Dock = 'Bottom'
$statusBar.Height = 30
$statusBar.BackColor = (Get-ThemeColors).StatusBar

$statusLabel = New-Object Windows.Forms.Label
$statusLabel.Text = (Get-LocalizedString "StatusReady")
$statusLabel.ForeColor = [System.Drawing.Color]::White
$statusLabel.Dock = 'Fill'
$statusLabel.TextAlign = 'MiddleLeft'
$statusBar.Controls.Add($statusLabel)
$form.Controls.Add($statusBar)

# События
$btnWU.Add_Click({ 
    $statusBox.AppendText("`r`n[INFO] Обращение к серверам Microsoft...`r`n")
    # Здесь логика запуска обновлений
})

$form.ShowDialog()