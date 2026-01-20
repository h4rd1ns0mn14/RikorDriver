# Driver Updater Rikor v 1.0
# Run as Administrator
# Usage: irm https://raw.githubusercontent.com/h4rd1ns0mn14/RikorDriver/refs/heads/main/RikorDriver.ps1 | iex
# Silent mode: .\RikorDriver.ps1 -Silent -Task "SmartUpdate"
param(
    [switch]$Silent,
    [string]$Task = "",
    [string]$Language = "ru",
    # [string]$ProxyAddress = "", # Removed proxy parameter
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
        AppTitle = "Rikor Driver Installer"
        BtnSmartUpdate = "Install & Update Drivers"
        BtnScan = "Scan for Missing Drivers"
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
        # ProxyConfigured = "Proxy configured:" # Removed
        # ProxyCleared = "Proxy settings cleared." # Removed
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
        # ProxyLabel = "Proxy Address:" # Removed
        ClassFilter = "Filter by Class:"
        ManufacturerFilter = "Filter by Manufacturer:"
        Apply = "Apply"
        Clear = "Clear"
        Close = "Close"
        Enable = "Enable"
        Disable = "Disable"
        Remove = "Remove Schedule"
        PSWU_Ensuring = "Ensuring PSWindowsUpdate module is available..."
        PSWU_Downloading = "Downloading PSWindowsUpdate module from PowerShell Gallery..."
        PSWU_Extracting = "Extracting PSWindowsUpdate module..."
        PSWU_Importing = "Importing PSWindowsUpdate module..."
        PSWU_Available = "PSWindowsUpdate module is available."
        PSWU_Failed = "Failed to ensure PSWindowsUpdate module availability."
        SmartUpdate_RikorDownload = "Attempting to download and install drivers from Rikor source..."
        SmartUpdate_RikorInstallFailed = "Rikor source installation failed. Trying Microsoft Update..."
        SmartUpdate_WUSearch = "Searching for driver updates on Microsoft Update..."
        SmartUpdate_WUInstall = "Installing updates from Microsoft Update..."
        SmartUpdate_WUFound = "Found {0} driver updates."
        SmartUpdate_WUNoUpdates = "No new driver updates found on Microsoft Update."
        SmartUpdate_WUInstallComplete = "Installation from Microsoft Update complete. A reboot may be required."
        SmartUpdate_WUInstallFailed = "Failed to check or install drivers from Microsoft Update."
        SmartUpdate_RikorDownloadFailed = "Failed to download or install from Rikor URL."
        SmartUpdate_ArchiveEmpty = "Downloaded file is too small or empty. Skipping Rikor source."
        WU_Probing = "Probing for driver updates (this may take a few moments)..."
        WU_InitiatingInstall = "Initiating installation of selected driver updates..."
        Scan_MissingDriversFound = "Found {0} missing driver(s):"
        Scan_NoMissingDrivers = "No missing driver updates found."
        Scan_ExportCSV = "Exporting missing drivers list to CSV..."
        Scan_Failed = "Failed to scan for missing drivers."
    }
    "ru" = @{
        AppTitle = "Установщик драйверов Rikor"
        BtnSmartUpdate = "Установить и обновить драйверы"
        BtnScan = "Сканировать недостающие драйверы" # Updated
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
        # ProxyConfigured = "Прокси настроен:" # Removed
        # ProxyCleared = "Настройки прокси сброшены." # Removed
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
        # ProxyLabel = "Адрес прокси:" # Removed
        ClassFilter = "Фильтр по классу:"
        ManufacturerFilter = "Фильтр по производителю:"
        Apply = "Применить"
        Clear = "Очистить"
        Close = "Закрыть"
        Enable = "Включить"
        Disable = "Отключить"
        Remove = "Удалить расписание"
        PSWU_Ensuring = "Проверка наличия модуля PSWindowsUpdate..."
        PSWU_Downloading = "Загрузка модуля PSWindowsUpdate с PowerShell Gallery..."
        PSWU_Extracting = "Распаковка модуля PSWindowsUpdate..."
        PSWU_Importing = "Импорт модуля PSWindowsUpdate..."
        PSWU_Available = "Модуль PSWindowsUpdate доступен."
        PSWU_Failed = "Не удалось обеспечить доступность модуля PSWindowsUpdate."
        SmartUpdate_RikorDownload = "Попытка загрузки и установки драйверов из источника Rikor..."
        SmartUpdate_RikorInstallFailed = "Не удалось установить из источника Rikor. Пробую Microsoft Update..."
        SmartUpdate_WUSearch = "Поиск обновлений драйверов в Центре обновления Windows..."
        SmartUpdate_WUInstall = "Установка обновлений из Центра обновления Windows..."
        SmartUpdate_WUFound = "Найдено {0} обновлений драйверов."
        SmartUpdate_WUNoUpdates = "Новых обновлений драйверов в Центре обновления Windows не найдено."
        SmartUpdate_WUInstallComplete = "Установка из Центра обновления Windows завершена. Может потребоваться перезагрузка."
        SmartUpdate_WUInstallFailed = "Не удалось проверить или установить драйверы из Центра обновления Windows."
        SmartUpdate_RikorDownloadFailed = "Не удалось загрузить или установить с URL-адреса Rikor."
        SmartUpdate_ArchiveEmpty = "Загруженный файл слишком мал или пуст. Пропускаю источник Rikor."
        WU_Probing = "Поиск обновлений драйверов (это может занять некоторое время)..."
        WU_InitiatingInstall = "Запуск установки выбранных обновлений драйверов..."
        Scan_MissingDriversFound = "Найдено {0} недостающих драйверов:"
        Scan_NoMissingDrivers = "Недостающие обновления драйверов не найдены."
        Scan_ExportCSV = "Экспорт списка недостающих драйверов в CSV..."
        Scan_Failed = "Не удалось отсканировать недостающие драйверы."
    }
}

# Get localized string
function Get-LocalizedString([string]$key, [array]$args = $null) {
    $lang = $global:CurrentLanguage
    if (-not $global:Languages.ContainsKey($lang)) { $lang = "en" }

    $string = if ($global:Languages[$lang].ContainsKey($key)) {
        $global:Languages[$lang][$key]
    } else {
        $global:Languages["en"][$key] # Fallback to English
    }

    if ($args) {
        return $string -f $args
    }
    return $string
}

$global:CurrentLanguage = $Language

# -------------------------
# Require Admin
# -------------------------
function Assert-AdminPrivilege {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
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
$LogBase = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller"
$HistoryFile = Join-Path $LogBase "UpdateHistory.json"
$SettingsFile = Join-Path $LogBase "Settings.json"
if (!(Test-Path $LogBase)) { New-Item -ItemType Directory -Path $LogBase -Force | Out-Null }

$global:CurrentJob = $null
$global:CurrentTaskLog = $null
# $global:ProxySettings = @{ Enabled = $false; Address = "" } # Removed
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
            # if ($settings.Proxy) { # Removed
            #    $global:ProxySettings.Enabled = $settings.Proxy.Enabled
            #    $global:ProxySettings.Address = $settings.Proxy.Address
            # } # Removed
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
        # Proxy = $global:ProxySettings # Removed
        Filters = $global:FilterSettings
        DarkMode = $global:DarkModeEnabled
    }
    $settings | ConvertTo-Json -Depth 3 | Set-Content -Path $SettingsFile -Encoding UTF8
}

Import-Settings

# Apply command-line parameters
# if ($ProxyAddress) { # Removed
#    $global:ProxySettings.Enabled = $true
#    $global:ProxySettings.Address = $ProxyAddress
# } # Removed
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
    if ($history.Count -gt 100) { $history = $history[0..99] }
    $history | ConvertTo-Json -Depth 3 | Set-Content -Path $HistoryFile -Encoding UTF8
}

function Get-UpdateHistory {
    if (Test-Path $HistoryFile) {
        try {
            $content = Get-Content -Path $HistoryFile -Raw
            if ($content) { return $content | ConvertFrom-Json }
        } catch {}
    }
    return @()
}

# -------------------------
# Network Proxy Support (Removed - keeping empty for structure consistency if needed later)
# -------------------------
# function Set-ProxySettings { ... } # Removed

# if ($global:ProxySettings.Enabled -and $global:ProxySettings.Address) { # Removed
#    Set-ProxySettings -ProxyAddr $global:ProxySettings.Address -Enable $true
# } # Removed

# -------------------------
# System Restore Point
# -------------------------
function New-RestorePoint {
    param([string]$Description = "Rikor Driver Installer Restore Point")
    try {
        Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
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
$ScheduledTaskName = "RikorDriverInstaller_ScheduledCheck"
function Get-ScheduledUpdateTask {
    try { return Get-ScheduledTask -TaskName $ScheduledTaskName -ErrorAction SilentlyContinue } catch { return $null }
}

function Set-ScheduledUpdate {
    param(
        [ValidateSet("Daily", "Weekly", "Monthly")]
        [string]$Frequency,
        [string]$Time = "03:00"
    )
    try {
        Remove-ScheduledUpdate
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { $scriptPath = $MyInvocation.PSCommandPath }
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -Silent -Task SmartUpdate"
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
        "SmartUpdate" {
            Write-SilentLog "Silent mode: Smart Update started."
            $zipUrl = "https://drive.google.com/uc?export=download&id=14_iaT8zdS800GpL76CSVb5vBQN7whZ8w"
            $tempDir = Join-Path $env:TEMP "RikorDriversSilent_$(Get-Date -Format 'yyyyMMddHHmmss')"
            $downloadSucceeded = $false

            try {
                New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                $zipPath = Join-Path $tempDir "drivers.zip"
                Write-SilentLog (Get-LocalizedString "SmartUpdate_RikorDownload")
                Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
                Write-SilentLog "Download completed."

                if ((Get-Item $zipPath).Length -lt 1024) { # Check for minimal size to avoid empty or error pages
                    Write-SilentLog (Get-LocalizedString "SmartUpdate_ArchiveEmpty")
                    throw "Downloaded file is empty or invalid."
                }

                $extractDir = Join-Path $tempDir "ExtractedDrivers"
                Write-SilentLog "Extracting archive to: $extractDir"
                Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

                $infFiles = Get-ChildItem -Path $extractDir -Recurse -Include *.inf
                if ($infFiles.Count -eq 0) {
                    Write-SilentLog "[WARNING] No .inf files found in the archive. Nothing to install from archive."
                } else {
                    Write-SilentLog "Found $($infFiles.Count) .inf files. Installing from Rikor pack..."
                    $successCount = 0; $failCount = 0
                    foreach ($inf in $infFiles) {
                        try {
                            & pnputil.exe /add-driver $inf.FullName /install /force 2>&1 | Out-Null
                            if ($LASTEXITCODE -eq 0) { $successCount++ } else { $failCount++ }
                        } catch { $failCount++ }
                    }
                    Write-SilentLog "Installation from Rikor pack complete: $successCount successful, $failCount failed."
                    $downloadSucceeded = $true
                }
            } catch {
                Write-SilentLog (Get-LocalizedString "SmartUpdate_RikorInstallFailed") + " $_"
            } finally {
                if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
            }

            if (-not $downloadSucceeded) {
                Write-SilentLog (Get-LocalizedString "SmartUpdate_WUSearch")
                try {
                    # Manual download and import of PSWindowsUpdate module for silent mode
                    $modulePath = Join-Path (Split-Path $MyInvocation.MyCommand.Definition) "Modules\PSWindowsUpdate"
                    if (-not (Test-Path $modulePath)) {
                        Write-SilentLog (Get-LocalizedString "PSWU_Downloading")
                        $moduleZipUrl = "https://www.powershellgallery.com/api/v2/package/PSWindowsUpdate" # Direct download link for the nupkg
                        $moduleZipPath = Join-Path $env:TEMP "PSWindowsUpdate.zip"

                        Invoke-WebRequest -Uri $moduleZipUrl -OutFile $moduleZipPath -UseBasicParsing -ErrorAction Stop
                        Write-SilentLog (Get-LocalizedString "PSWU_Extracting")

                        # Nupkg is actually a zip file. Extract it.
                        Expand-Archive -Path $moduleZipPath -DestinationPath $modulePath -Force -ErrorAction Stop

                        # Nupkg contains a folder structure like PSWindowsUpdate\5.0.0.1\PSWindowsUpdate.psd1
                        # We need the actual module folder, which is typically the version folder
                        $versionFolder = (Get-ChildItem -Path $modulePath -Directory | Sort-Object Name -Descending | Select-Object -First 1).FullName
                        if ($versionFolder) {
                            Move-Item -Path "$versionFolder\*" -Destination $modulePath -Force
                            Remove-Item -Path $versionFolder -Recurse -Force
                        }
                        Remove-Item $moduleZipPath -Force -ErrorAction SilentlyContinue
                        Write-SilentLog (Get-LocalizedString "PSWU_Available")
                    }
                    Import-Module PSWindowsUpdate -Force -ErrorAction Stop

                    Write-SilentLog (Get-LocalizedString "WU_Probing")
                    $updates = Get-WindowsUpdate -Driver -ErrorAction Stop

                    if ($null -eq $updates -or $updates.Count -eq 0) {
                        Write-SilentLog (Get-LocalizedString "SmartUpdate_WUNoUpdates")
                    } else {
                        Write-SilentLog (Get-LocalizedString "SmartUpdate_WUFound", @($updates.Count))
                        Write-SilentLog (Get-LocalizedString "WU_InitiatingInstall")
                        $updates | ForEach-Object { Write-SilentLog " -> $($_.Title)" }
                        Install-WindowsUpdate -Driver -AcceptAll -IgnoreReboot -ErrorAction Stop
                        Write-SilentLog (Get-LocalizedString "SmartUpdate_WUInstallComplete")
                    }
                } catch {
                    Write-SilentLog (Get-LocalizedString "SmartUpdate_WUInstallFailed") + " $_"
                    Add-HistoryEntry -TaskName "SmartUpdate" -Status "Failed" -Details (Get-LocalizedString "SmartUpdate_WUInstallFailed") + " $_"
                }
            }
        }
        "ScanDrivers" { # Updated for silent mode as well
            Write-SilentLog "Silent mode: Scanning for missing drivers..."
            try {
                # Manual download and import of PSWindowsUpdate module for silent mode
                $modulePath = Join-Path (Split-Path $MyInvocation.MyCommand.Definition) "Modules\PSWindowsUpdate"
                if (-not (Test-Path $modulePath)) {
                    Write-SilentLog (Get-LocalizedString "PSWU_Downloading")
                    $moduleZipUrl = "https://www.powershellgallery.com/api/v2/package/PSWindowsUpdate" # Direct download link for the nupkg
                    $moduleZipPath = Join-Path $env:TEMP "PSWindowsUpdate.zip"

                    Invoke-WebRequest -Uri $moduleZipUrl -OutFile $moduleZipPath -UseBasicParsing -ErrorAction Stop
                    Write-SilentLog (Get-LocalizedString "PSWU_Extracting")

                    Expand-Archive -Path $moduleZipPath -DestinationPath $modulePath -Force -ErrorAction Stop
                    $versionFolder = (Get-ChildItem -Path $modulePath -Directory | Sort-Object Name -Descending | Select-Object -First 1).FullName
                    if ($versionFolder) {
                        Move-Item -Path "$versionFolder\*" -Destination $modulePath -Force
                        Remove-Item -Path $versionFolder -Recurse -Force
                    }
                    Remove-Item $moduleZipPath -Force -ErrorAction SilentlyContinue
                    Write-SilentLog (Get-LocalizedString "PSWU_Available")
                }
                Import-Module PSWindowsUpdate -Force -ErrorAction Stop

                Write-SilentLog (Get-LocalizedString "WU_Probing")
                $missingDrivers = Get-WindowsUpdate -Driver -NotInstalled -ErrorAction Stop

                $filteredDrivers = $missingDrivers
                # Apply filters (if needed for silent mode)
                # Note: Silent mode filter settings come from command-line parameters, not global:FilterSettings
                if ($FilterClass) {
                    $filteredDrivers = $filteredDrivers | Where-Object { $_.Categories -join ', ' -like "*$($FilterClass)*" }
                }
                if ($FilterManufacturer) {
                    $filteredDrivers = $filteredDrivers | Where-Object { $_.Manufacturer -like "*$($FilterManufacturer)*" }
                }

                if ($filteredDrivers.Count -eq 0) {
                    Write-SilentLog (Get-LocalizedString "Scan_NoMissingDrivers")
                } else {
                    Write-SilentLog (Get-LocalizedString "Scan_MissingDriversFound", @($filteredDrivers.Count))
                    foreach ($driver in $filteredDrivers) {
                        Write-SilentLog "  - Title: $($driver.Title) (Manufacturer: $($driver.Manufacturer), Version: $($driver.Version))"
                    }

                    Write-SilentLog (Get-LocalizedString "Scan_ExportCSV")
                    $csvPath = Join-Path (Split-Path $logFile) "MissingDrivers_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
                    $filteredDrivers | Select-Object Title, Manufacturer, @{Name='Categories';Expression={ $_.Categories -join ', ' }}, Version, LastDeploymentChangeTime | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                    Write-SilentLog "Exported to: $csvPath"
                }
            } catch {
                Write-SilentLog "[ERROR] " + (Get-LocalizedString "Scan_Failed") + ": $_"
            }
        }
        "BackupDrivers" {
            $dest = $innerArgs[0]
            Write-SilentLog "Backing up drivers to: $dest"
            try {
                if (!(Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }
                & dism.exe /online /export-driver /destination:$dest 2>&1 | Out-Null
                $exportedCount = (Get-ChildItem -Path $dest -Recurse -Directory -ErrorAction SilentlyContinue).Count
                Write-SilentLog "Backup completed: $exportedCount driver package(s) exported."
            } catch {
                Write-SilentLog "[ERROR] Backup failed: $_"
            }
        }
        "InstallDrivers" {
            $folder = $innerArgs[0]
            Write-SilentLog "Installing drivers from: $folder"
            try {
                if (-not (Test-Path $folder)) { Write-SilentLog "[ERROR] Folder not found: $folder"; return }
                $infFiles = Get-ChildItem -Path $folder -Recurse -Include *.inf -ErrorAction SilentlyContinue
                if ($infFiles.Count -eq 0) { Write-SilentLog "[ERROR] No .inf driver files found in folder"; return }
                $successCount = 0; $failCount = 0
                foreach ($inf in $infFiles) {
                    try {
                        & pnputil.exe /add-driver $inf.FullName /install /force 2>&1 | Out-Null
                        if ($LASTEXITCODE -eq 0) { $successCount++ } else { $failCount++ }
                    } catch { $failCount++ }
                }
                Write-SilentLog "Installation complete: $successCount successful, $failCount failed."
            } catch {
                Write-SilentLog "[ERROR] Installation failed: $_"
            }
        }
        default {
            Write-SilentLog "Unknown task: $Task"
            Add-HistoryEntry -TaskName $Task -Status "Failed" -Details "Unknown task"
        }
    }
    Write-SilentLog "Silent mode completed."
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
    Dark = @{ Background = [System.Drawing.Color]::FromArgb(18, 18, 18); Surface = [System.Drawing.Color]::FromArgb(30, 30, 30); SurfaceHover = [System.Drawing.Color]::FromArgb(45, 45, 45); Primary = [System.Drawing.Color]::FromArgb(156, 39, 176); PrimaryHover = [System.Drawing.Color]::FromArgb(186, 74, 206); Secondary = [System.Drawing.Color]::FromArgb(66, 66, 66); Text = [System.Drawing.Color]::FromArgb(240, 240, 240); TextSecondary = [System.Drawing.Color]::FromArgb(170, 170, 170); Border = [System.Drawing.Color]::FromArgb(60, 60, 60); Success = [System.Drawing.Color]::FromArgb(76, 175, 80); Warning = [System.Drawing.Color]::FromArgb(255, 152, 0); Error = [System.Drawing.Color]::FromArgb(244, 67, 54); MenuBar = [System.Drawing.Color]::FromArgb(25, 25, 25); StatusBar = [System.Drawing.Color]::FromArgb(126, 34, 152) }
    Light = @{ Background = [System.Drawing.Color]::FromArgb(255, 255, 255); Surface = [System.Drawing.Color]::FromArgb(245, 247, 250); SurfaceHover = [System.Drawing.Color]::FromArgb(235, 238, 242); Primary = [System.Drawing.Color]::FromArgb(156, 39, 176); PrimaryHover = [System.Drawing.Color]::FromArgb(186, 74, 206); Secondary = [System.Drawing.Color]::FromArgb(224, 224, 224); Text = [System.Drawing.Color]::FromArgb(33, 33, 33); TextSecondary = [System.Drawing.Color]::FromArgb(117, 117, 117); Border = [System.Drawing.Color]::FromArgb(224, 224, 224); Success = [System.Drawing.Color]::FromArgb(67, 160, 71); Warning = [System.Drawing.Color]::FromArgb(251, 1, 0); Error = [System.Drawing.Color]::FromArgb(229, 57, 53); MenuBar = [System.Drawing.Color]::FromArgb(255, 255, 255); StatusBar = [System.Drawing.Color]::FromArgb(156, 39, 176) }
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
$menuFile = New-Object Windows.Forms.ToolStripMenuItem("&File")
$menuOpenLogs = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnOpenLogs"))
$menuOpenLogs.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::L
$menuExit = New-Object Windows.Forms.ToolStripMenuItem("Exit")
$menuExit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$menuFile.DropDownItems.AddRange(@($menuOpenLogs, (New-Object Windows.Forms.ToolStripSeparator), $menuExit))
$menuActions = New-Object Windows.Forms.ToolStripMenuItem("&Actions")
$menuSmartUpdate = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnSmartUpdate"))
$menuSmartUpdate.ShortcutKeys = [System.Windows.Forms.Keys]::F5
$menuScan = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnScan"))
$menuScan.ShortcutKeys = [System.Windows.Forms.Keys]::F7
$menuBackup = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnBackup"))
$menuBackup.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::B
$menuInstall = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnInstall"))
$menuInstall.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::I
$menuCancel = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnCancel"))
$menuCancel.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Q
$menuActions.DropDownItems.AddRange(@($menuSmartUpdate, $menuScan, (New-Object Windows.Forms.ToolStripSeparator), $menuBackup, $menuInstall, (New-Object Windows.Forms.ToolStripSeparator), $menuCancel))
$menuTools = New-Object Windows.Forms.ToolStripMenuItem("&Tools")
$menuRestorePoint = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnRestorePoint"))
$menuSchedule = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnSchedule"))
$menuFilters = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnFilters"))
$menuHistory = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnHistory"))
$menuTools.DropDownItems.AddRange(@($menuRestorePoint, $menuSchedule, $menuFilters, (New-Object Windows.Forms.ToolStripSeparator), $menuHistory))
$menuView = New-Object Windows.Forms.ToolStripMenuItem("&View")
$menuToggleTheme = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnDarkMode"))
$menuLanguage = New-Object Windows.Forms.ToolStripMenuItem("Language")
$langItems = @("English (en)", "Русский (ru)")
$langCodes = @("en", "ru")
for ($i = 0; $i -lt $langItems.Count; $i++) {
    $langMenuItem = New-Object Windows.Forms.ToolStripMenuItem($langItems[$i])
    $langMenuItem.Tag = $langCodes[$i]
    if ($langCodes[$i] -eq $global:CurrentLanguage) { $langMenuItem.Checked = $true }
    $menuLanguage.DropDownItems.Add($langMenuItem) | Out-Null
}
$menuView.DropDownItems.AddRange(@($menuToggleTheme, (New-Object Windows.Forms.ToolStripSeparator), $menuLanguage))
$menuSettingsTop = New-Object Windows.Forms.ToolStripMenuItem("&Settings")
$menuStrip.Items.AddRange(@($menuFile, $menuActions, $menuTools, $menuView, $menuSettingsTop))
$form.MainMenuStrip = $menuStrip

# -------------------------
# Toolbar Panel
# -------------------------
$toolbarPanel = New-Object Windows.Forms.Panel
$toolbarPanel.Dock = 'Top'
$toolbarPanel.Height = 56
$buttonContainer = New-Object Windows.Forms.FlowLayoutPanel
$buttonContainer.Dock = 'Fill'
$buttonContainer.FlowDirection = 'LeftToRight'
$buttonContainer.WrapContents = $false
$buttonContainer.AutoSize = $false
$buttonContainer.Padding = '0,8,0,8'
$toolbarPanel.Controls.Add($buttonContainer)

function Update-ButtonContainerPadding {
    # Adjusted widths for the new buttons
    $totalButtonWidth = 220 + 155 + 120 + 140 + 110 + (4 * 12)  # buttons + gaps
    $availableWidth = $toolbarPanel.ClientSize.Width
    $leftPadding = [Math]::Max(0, [int](($availableWidth - $totalButtonWidth) / 2))
    $buttonContainer.Padding = "$leftPadding,8,0,8"
}

function New-RoundedRegion {
    param([int]$Width, [int]$Height, [int]$Radius = 8)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $r = $Radius
    $w = $Width
    $h = $Height
    $path.AddArc(0, 0, $r * 2, $r * 2, 180, 90)
    $path.AddArc($w - $r * 2, 0, $r * 2, $r * 2, 270, 90)
    $path.AddArc($w - $r * 2, $h - $r * 2, $r * 2, $r * 2, 0, 90)
    $path.AddArc(0, $h - $r * 2, $r * 2, $r * 2, 90, 90)
    $path.CloseFigure()
    return New-Object System.Drawing.Region($path)
}

function New-ModernButton {
    param(
        [string]$Text,
        [int]$Width = 130,
        [bool]$Primary = $false
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Width = $Width
    $btn.Height = 38
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Font = New-Object Drawing.Font("Segoe UI Semibold", 9)
    $btn.Tag = $Primary
    $btn.TextAlign = 'MiddleCenter'
    $btn.Region = New-RoundedRegion -Width $Width -Height 38 -Radius 6
    return $btn
}

$btnSmartUpdate = New-ModernButton -Text (Get-LocalizedString "BtnSmartUpdate") -Width 220 -Primary $true
$btnSmartUpdate.Margin = '0,0,12,0'
$btnScan = New-ModernButton -Text (Get-LocalizedString "BtnScan") -Width 155
$btnScan.Margin = '0,0,12,0'
$btnBackup = New-ModernButton -Text (Get-LocalizedString "BtnBackup") -Width 120
$btnBackup.Margin = '0,0,12,0'
$btnInstall = New-ModernButton -Text (Get-LocalizedString "BtnInstall") -Width 140
$btnInstall.Margin = '0,0,12,0'
$btnCancel = New-ModernButton -Text (Get-LocalizedString "BtnCancel") -Width 110
$buttonContainer.Controls.AddRange(@($btnSmartUpdate, $btnScan, $btnBackup, $btnInstall, $btnCancel))

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
$versionLabel.Text = "v3.1  "
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
$form.Controls.Add($contentPanel)    # Fill - takes remaining space
$form.Controls.Add($statusBar)       # Bottom
$form.Controls.Add($toolbarSeparator) # Top - appears below toolbar
$form.Controls.Add($toolbarPanel)    # Top - appears below menu
$form.Controls.Add($menuStrip)       # Top - appears at very top
$form.MainMenuStrip = $menuStrip

$statusBorderPanel = New-Object Windows.Forms.Panel
$statusBorderPanel.Dock = 'Fill'
$statusBorderPanel.Padding = '1,1,1,1'
$contentPanel.Controls.Add($statusBorderPanel)

$status = New-Object Windows.Forms.RichTextBox
$status.Multiline = $true
$status.ReadOnly = $true
$status.Dock = 'Fill'
$status.ScrollBars = 'Vertical'
$status.BorderStyle = 'None'
$status.Font = New-Object Drawing.Font("Cascadia Code, Consolas, Courier New", 9.5)
$statusBorderPanel.Controls.Add($status)

$progressPanel = New-Object Windows.Forms.Panel
$progressPanel.Dock = 'Bottom'
$progressPanel.Height = 36
$progressPanel.Padding = '0,8,0,8'
$contentPanel.Controls.Add($progressPanel)

$progressBorderPanel = New-Object Windows.Forms.Panel
$progressBorderPanel.Dock = 'Fill'
$progressBorderPanel.Padding = '1,1,1,1'
$progressPanel.Controls.Add($progressBorderPanel)

$progress = New-Object Windows.Forms.ProgressBar
$progress.Dock = 'Fill'
$progress.Style = 'Continuous'
$progress.Value = 0
$progressBorderPanel.Controls.Add($progress)

$headerLabel = New-Object Windows.Forms.Label
$headerLabel.Text = "Output Console"
$headerLabel.Dock = 'Top'
$headerLabel.Height = 26
$headerLabel.Font = New-Object Drawing.Font("Segoe UI Semibold", 9.5)
$headerLabel.TextAlign = 'MiddleLeft'
$contentPanel.Controls.Add($headerLabel)

$cmbLang = New-Object Windows.Forms.ComboBox
$cmbLang.Items.AddRange(@("en", "es", "fr", "de", "pt", "it", "ru"))
$cmbLang.SelectedItem = $global:CurrentLanguage

$driversGrid = New-Object Windows.Forms.DataGridView
$driversGrid.ReadOnly = $true
$driversGrid.AllowUserToAddRows = $false
$driversGrid.AllowUserToDeleteRows = false
$driversGrid.Height = 60
$driversGrid.Visible = $false

$timer = New-Object Windows.Forms.Timer
$timer.Interval = 500

# -------------------------
# Ensure PSWindowsUpdate Module
# (Replaces Install-PSWindowsUpdateModule to avoid NuGet/PowerShellGet issues)
# -------------------------
function Ensure-PSWindowsUpdateModule {
    param($formRef, $statusRef)

    $moduleName = "PSWindowsUpdate"
    $modulePath = Join-Path (Split-Path $MyInvocation.MyCommand.Definition) "Modules\$moduleName"

    # Check if module is already imported or available
    if (Get-Module -Name $moduleName -ErrorAction SilentlyContinue) {
        Add-StatusUI $formRef $statusRef "[INFO] $moduleName module is already loaded."
        return $true
    }
    if (Test-Path "$modulePath\$moduleName.psd1") {
        Add-StatusUI $formRef $statusRef "[INFO] $moduleName module found locally. Importing..."
        try {
            Import-Module $modulePath -Force -ErrorAction Stop
            Add-StatusUI $formRef $statusRef (Get-LocalizedString "PSWU_Available")
            return $true
        } catch {
            Add-StatusUI $formRef $statusRef "[ERROR] Failed to import $moduleName from local path: $_"
            # Fall through to download if import fails
        }
    }

    Add-StatusUI $formRef $statusRef (Get-LocalizedString "PSWU_Ensuring")

    try {
        # Attempt to find it in default module paths
        $defaultModulePath = (Get-Module -ListAvailable -Name $moduleName).Path
        if ($defaultModulePath) {
            Add-StatusUI $formRef $statusRef "[INFO] $moduleName found in system path: $defaultModulePath. Importing..."
            Import-Module $moduleName -Force -ErrorAction Stop
            Add-StatusUI $formRef $statusRef (Get-LocalizedString "PSWU_Available")
            return $true
        }

        # If not found, download and extract
        Add-StatusUI $formRef $statusRef (Get-LocalizedString "PSWU_Downloading")
        $moduleZipUrl = "https://www.powershellgallery.com/api/v2/package/$moduleName" # Direct download link for the nupkg
        $moduleZipPath = Join-Path $env:TEMP "$moduleName.zip"

        # Create Modules directory if it doesn't exist
        if (-not (Test-Path $modulePath)) { New-Item -ItemType Directory -Path $modulePath -Force | Out-Null }

        Invoke-WebRequest -Uri $moduleZipUrl -OutFile $moduleZipPath -UseBasicParsing -ErrorAction Stop
        Add-StatusUI $formRef $statusRef (Get-LocalizedString "PSWU_Extracting")

        # Nupkg is actually a zip file. Extract it.
        Expand-Archive -Path $moduleZipPath -DestinationPath $modulePath -Force -ErrorAction Stop

        # Nupkg contains a folder structure like PSWindowsUpdate\5.0.0.1\PSWindowsUpdate.psd1
        # We need to move the actual module content up one level
        $versionFolder = (Get-ChildItem -Path $modulePath -Directory | Sort-Object Name -Descending | Select-Object -First 1).FullName
        if ($versionFolder) {
            Move-Item -Path "$versionFolder\*" -Destination $modulePath -Force
            Remove-Item -Path $versionFolder -Recurse -Force
        }
        Remove-Item $moduleZipPath -Force -ErrorAction SilentlyContinue

        Add-StatusUI $formRef $statusRef (Get-LocalizedString "PSWU_Importing")
        Import-Module $modulePath -Force -ErrorAction Stop
        Add-StatusUI $formRef $statusRef (Get-LocalizedString "PSWU_Available")
        return $true
    } catch {
        Add-StatusUI $formRef $statusRef "[ERROR] $(Get-LocalizedString 'PSWU_Failed'): $_"
        return $false
    }
}

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

    # Ensure PSWindowsUpdate module is available before starting tasks that depend on it
    if ($Name -eq "SmartUpdate" -or $Name -eq "ScanDrivers") {
        if (-not (Ensure-PSWindowsUpdateModule -formRef $form -statusRef $status)) {
            return $null # Stop if module cannot be ensured
        }
    }

    $log = New-TaskLog $Name
    $global:CurrentTaskLog = $log
    Add-StatusUI $form $status "$(Get-LocalizedString 'StartingTask') $Name"

    # Pass global variables explicitly to the job's script block
    $job = Start-Job -Name $Name -ScriptBlock {
        param($taskName, $logPath, $innerArgs, $jobLanguages, $jobCurrentLanguage, $jobFilterSettings, $scriptDir) # Added scriptDir

        # Redefine Get-LocalizedString within the job's scope
        function Get-LocalizedString([string]$key, [array]$args = $null) {
            $lang = $jobCurrentLanguage
            if (-not $jobLanguages.ContainsKey($lang)) { $lang = "en" }

            $string = if ($jobLanguages[$lang].ContainsKey($key)) {
                $jobLanguages[$lang][$key]
            } else {
                $jobLanguages["en"][$key] # Fallback to English
            }

            if ($args) {
                return $string -f $args
            }
            return $string
        }

        function L($m) {
            $t = (Get-Date).ToString("s")
            Add-Content -Path $logPath -Value ("$t - $m")
        }

        try {
            # Import PSWindowsUpdate module inside the job, as it's a separate session
            # We assume Ensure-PSWindowsUpdateModule has already made it available locally
            if ($taskName -eq "SmartUpdate" -or $taskName -eq "ScanDrivers") {
                try {
                    $moduleName = "PSWindowsUpdate"
                    $localModulePath = Join-Path $scriptDir "Modules\$moduleName"
                    if (Test-Path "$localModulePath\$moduleName.psd1") {
                         Import-Module $localModulePath -Force -ErrorAction Stop
                         L "[INFO] PSWindowsUpdate module imported successfully within job."
                    } else {
                         # Fallback to system-wide import if local path fails (should not happen if Ensure worked)
                         Import-Module $moduleName -Force -ErrorAction Stop
                         L "[INFO] PSWindowsUpdate module imported from system path within job."
                    }
                } catch {
                    L "[ERROR] Failed to import PSWindowsUpdate module within job: $_"
                    throw "Module import failed in background job."
                }
            }

            switch ($taskName) {
                "SmartUpdate" {
                    L (Get-LocalizedString "SmartUpdate_RikorDownload")
                    $zipUrl = "https://drive.google.com/uc?export=download&id=14_iaT8zdS800GpL76CSVb5vBQN7whZ8w"
                    $tempDir = Join-Path $env:TEMP "RikorDrivers_$(Get-Date -Format 'yyyyMMddHHmmss')"
                    $downloadSucceeded = $false

                    try {
                        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                        $zipPath = Join-Path $tempDir "drivers.zip"
                        L "URL: $zipUrl"
                        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
                        L "Download completed successfully."

                        if ((Get-Item $zipPath).Length -lt 1024) { # Check for minimal size to avoid empty or error pages
                            L (Get-LocalizedString "SmartUpdate_ArchiveEmpty")
                            throw "Downloaded file is empty or invalid."
                        }

                        $extractDir = Join-Path $tempDir "ExtractedDrivers"
                        L "Extracting archive to: $extractDir"
                        Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

                        $infFiles = Get-ChildItem -Path $extractDir -Recurse -Include *.inf
                        if ($infFiles.Count -eq 0) {
                            L "[WARNING] No .inf files found in the archive. Nothing to install from archive."
                        } else {
                            L "Found $($infFiles.Count) .inf files. Installing from Rikor pack..."
                            $successCount = 0; $failCount = 0
                            foreach ($inf in $infFiles) {
                                L "Installing $($inf.Name)..."
                                try {
                                    & pnputil.exe /add-driver $inf.FullName /install /force 2>&1 | Out-Null
                                    if ($LASTEXITCODE -eq 0) { $successCount++ } else { $failCount++; L " -> Failed (Code: $LASTEXITCODE)" }
                                } catch { $failCount++; L " -> Failed (Exception: $_)" }
                            }
                            L "Installation from Rikor pack complete: $successCount successful, $failCount failed."
                            $downloadSucceeded = $true
                        }
                    } catch {
                        L (Get-LocalizedString "SmartUpdate_RikorDownloadFailed") + " $_. "
                        L (Get-LocalizedString "SmartUpdate_RikorInstallFailed")
                    } finally {
                        if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
                    }
                    L " " # Spacer
                    if (-not $downloadSucceeded) {
                         L (Get-LocalizedString "SmartUpdate_WUSearch")
                    } else {
                         L "Rikor drivers installed, now checking for additional updates from Microsoft Update."
                    }

                    try {
                        L (Get-LocalizedString "WU_Probing") # Added this log
                        $driverUpdates = Get-WindowsUpdate -Driver -ErrorAction Stop

                        if ($null -eq $driverUpdates -or $driverUpdates.Count -eq 0) {
                            L (Get-LocalizedString "SmartUpdate_WUNoUpdates")
                        } else {
                            L (Get-LocalizedString "SmartUpdate_WUFound", @($driverUpdates.Count))
                            L (Get-LocalizedString "WU_InitiatingInstall") # Added this log
                            $driverUpdates | ForEach-Object { L " -> $($_.Title) (Manufacturer: $($_.Manufacturer))" } # Improved logging for what's being installed

                            # Install drivers found
                            Install-WindowsUpdate -Driver -AcceptAll -IgnoreReboot -ErrorAction Stop
                            L (Get-LocalizedString "SmartUpdate_WUInstallComplete")
                        }
                    } catch {
                        L "[ERROR] " + (Get-LocalizedString "SmartUpdate_WUInstallFailed") + ": $_"
                    }
                }
                "ScanDrivers" { # Updated for scanning missing drivers
                    L "Scanning for missing drivers..."
                    try {
                        L (Get-LocalizedString "WU_Probing")
                        $missingDrivers = Get-WindowsUpdate -Driver -NotInstalled -ErrorAction Stop

                        $filteredDrivers = $missingDrivers
                        # Apply filters
                        if ($jobFilterSettings.Class) {
                            L "Applying class filter: $($jobFilterSettings.Class)"
                            $filteredDrivers = $filteredDrivers | Where-Object { $_.Categories -join ', ' -like "*$($jobFilterSettings.Class)*" }
                        }
                        if ($jobFilterSettings.Manufacturer) {
                            L "Applying manufacturer filter: $($jobFilterSettings.Manufacturer)"
                            $filteredDrivers = $filteredDrivers | Where-Object { $_.Manufacturer -like "*$($jobFilterSettings.Manufacturer)*" }
                        }

                        if ($filteredDrivers.Count -eq 0) {
                            L (Get-LocalizedString "Scan_NoMissingDrivers")
                        } else {
                            L (Get-LocalizedString "Scan_MissingDriversFound", @($filteredDrivers.Count))
                            L "----------------------------------------------------"
                            $filteredDrivers | ForEach-Object {
                                L "  - Title: $($_.Title)"
                                L "    Manufacturer: $($_.Manufacturer)"
                                L "    Version: $($_.Version)"
                                L "    Categories: $($_.Categories -join ', ')"
                                L "----------------------------------------------------"
                            }

                            L (Get-LocalizedString "Scan_ExportCSV")
                            $csvPath = Join-Path (Split-Path $logPath) "MissingDrivers_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
                            $filteredDrivers | Select-Object Title, Manufacturer, @{Name='Categories';Expression={ $_.Categories -join ', ' }}, Version, LastDeploymentChangeTime | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                            L "Exported to: $csvPath"
                        }
                    } catch {
                        L "[ERROR] " + (Get-LocalizedString "Scan_Failed") + ": $_"
                    }
                }
                "BackupDrivers" {
                    $dest = $innerArgs[0]
                    L "Backing up drivers to: $dest"
                    try {
                        if (!(Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }
                        L "Exporting drivers (this may take several minutes)..."
                        & dism.exe /online /export-driver /destination:$dest 2>&1 | Out-Null
                        $exportedCount = (Get-ChildItem -Path $dest -Recurse -Directory -ErrorAction SilentlyContinue).Count
                        L "Backup completed: $exportedCount driver package(s) exported."
                    } catch {
                        L "[ERROR] Backup failed: $_"
                    }
                }
                "InstallDrivers" {
                    $folder = $innerArgs[0]
                    L "Installing drivers from: $folder"
                    try {
                        if (-not (Test-Path $folder)) { L "[ERROR] Folder not found: $folder"; return }
                        $infFiles = Get-ChildItem -Path $folder -Recurse -Include *.inf -ErrorAction SilentlyContinue
                        if ($infFiles.Count -eq 0) { L "[ERROR] No .inf driver files found in folder"; return }
                        L "Found $($infFiles.Count) driver file(s). Installing..."
                        $successCount = 0; $failCount = 0
                        foreach ($inf in $infFiles) {
                            L "Installing $($inf.Name)..."
                            try {
                                & pnputil.exe /add-driver $inf.FullName /install /force 2>&1 | Out-Null
                                if ($LASTEXITCODE -eq 0) { $successCount++ } else { $failCount++ }
                            } catch { $failCount++ }
                        }
                        L "Installation complete: $successCount successful, $failCount failed."
                    } catch {
                        L "[ERROR] Installation failed: $_"
                    }
                }
                default {
                    L "ERROR: Unknown task name: $taskName"
                }
            }
        } catch {
            L "FATAL ERROR in job: $_"
        } finally {
            L "Completed"
        }
    } -ArgumentList $Name, $log, $TaskArgs, $global:Languages, $global:CurrentLanguage, $global:FilterSettings, $PSScriptRoot # Pass script directory

    $global:CurrentJob = $job
    $timer.Start()
    return $job
}

# Timer tick event to update UI
$timer.Add_Tick({
    try {
        if ($global:CurrentTaskLog -and (Test-Path $global:CurrentTaskLog)) {
            $lines = Get-Content -Path $global:CurrentTaskLog -Tail 100 -ErrorAction SilentlyContinue
            if ($lines) {
                $text = ($lines -join "`r`n")
                $form.Invoke([System.Windows.Forms.MethodInvoker]{
                    if($status.Text.Length -ne $text.Length) { # Only update if text content has changed
                        $status.Text = $text
                        $status.ScrollToCaret()
                    }
                })
            }
        }
        # Update progress bar based on log content (heuristic)
        if ($global:CurrentTaskLog -and (Test-Path $global:CurrentTaskLog)) {
            $content = Get-Content -Path $global:CurrentTaskLog -Tail 100 -ErrorAction SilentlyContinue -Raw
            $p = $progress.Value # Keep current progress if no new markers found

            # SmartUpdate progress
            if ($content -match "Smart Update started") { $p = 1 }
            if ($content -match "Attempting to download") { $p = 5 }
            if ($content -match "Download completed successfully") { $p = 20 }
            if ($content -match "Extracting archive") { $p = 25 }
            if ($content -match "Installing from Rikor pack") { $p = 40 }
            if ($content -match "Installation from Rikor pack complete") { $p = 55 }
            if ($content -match "Rikor drivers installed, now checking for additional updates from Microsoft Update." -or $content -match "Rikor source installation failed. Trying Microsoft Update...") { $p = 60 }
            if ($content -match "Ensuring PSWindowsUpdate module is available" -or $content -match "Downloading PSWindowsUpdate module") { $p = 62 }
            if ($content -match "Extracting PSWindowsUpdate module") { $p = 65 }
            if ($content -match "Importing PSWindowsUpdate module") { $p = 67 }
            if ($content -match "PSWindowsUpdate module is available") { $p = 70 }
            if ($content -match "Searching for driver updates on Microsoft Update" -or $content -match "Probing for driver updates") { $p = 75 }
            if ($content -match "Found \d+ driver updates") { $p = 80 }
            if ($content -match "Initiating installation of selected driver updates" -or $content -match "Installing updates from Microsoft Update") { $p = 85 }
            if ($content -match "Installation from Microsoft Update complete") { $p = 95 }

            # ScanDrivers progress
            if ($content -match "Scanning for missing drivers") { $p = 10 }
            if ($content -match "Ensuring PSWindowsUpdate module is available") { $p = 15 }
            if ($content -match "Downloading PSWindowsUpdate module") { $p = 20 }
            if ($content -match "Extracting PSWindowsUpdate module") { $p = 25 }
            if ($content -match "Importing PSWindowsUpdate module") { $p = 30 }
            if ($content -match "PSWindowsUpdate module is available") { $p = 35 }
            if ($content -match "Probing for driver updates") { $p = 40 }
            if ($content -match "Found \d+ missing driver\(s\)") { $p = 80 }
            if ($content -match "Exporting missing drivers list to CSV") { $p = 90 }

            if ($content -match "Completed") { $p = 100 }

            # For general InstallDrivers
            if ($content -match "Found \d+ driver file\(s\)\. Installing") { $p = 30 }

            $progress.Value = [Math]::Min(100, [int]$p)
        }


        if ($null -ne $global:CurrentJob) {
            $jobState = (Get-Job -Id $global:CurrentJob.Id -ErrorAction SilentlyContinue).State
            if ($jobState -in @("Completed", "Failed", "Stopped")) {
                Start-Sleep -Milliseconds 200 # Give a moment for final logs to be written
                Add-HistoryEntry -TaskName $global:CurrentJob.Name -Status $jobState -Details "Task finished."
                $form.Invoke([System.Windows.Forms.MethodInvoker]{
                    $status.AppendText("`r`n=== $(Get-LocalizedString 'TaskFinished') $jobState ===`r`n")
                    $statusLabel.Text = "  Task $jobState"
                    $progress.Value = 100 # Ensure progress is 100% on completion
                })
                Remove-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
                $global:CurrentJob = $null
                $timer.Stop()
            }
        }
    } catch {
        # ignore UI timer exceptions to prevent crashing the UI
    }
})

# -------------------------
# Theme Functions
# -------------------------
function Set-Theme {
    param([bool]$Dark)
    $global:DarkModeEnabled = $Dark
    $colors = Get-ThemeColors
    $form.BackColor = $colors.Background
    $menuStrip.BackColor = $colors.MenuBar
    $menuStrip.ForeColor = $colors.Text
    foreach ($item in $menuStrip.Items) {
        $item.BackColor = $colors.MenuBar
        $item.ForeColor = $colors.Text
    }
    $toolbarPanel.BackColor = $colors.Surface
    $buttonContainer.BackColor = $colors.Surface
    $toolbarSeparator.BackColor = $colors.Border
    foreach ($ctrl in $buttonContainer.Controls) {
        if ($ctrl -is [System.Windows.Forms.Button]) {
            if ($ctrl.Tag -eq $true) {
                $ctrl.BackColor = $colors.Primary
                $ctrl.ForeColor = [System.Drawing.Color]::White
                $ctrl.FlatAppearance.MouseOverBackColor = $colors.PrimaryHover
            } else {
                $ctrl.BackColor = $colors.Secondary
                $ctrl.ForeColor = $colors.Text
                $ctrl.FlatAppearance.MouseOverBackColor = $colors.SurfaceHover
            }
        }
    }
    $contentPanel.BackColor = $colors.Background
    $headerLabel.BackColor = $colors.Background
    $headerLabel.ForeColor = $colors.TextSecondary
    $statusBorderPanel.BackColor = $colors.Border
    $status.BackColor = $colors.Surface
    $status.ForeColor = $colors.Text
    $progressPanel.BackColor = $colors.Background
    $progressBorderPanel.BackColor = $colors.Border
    $statusBar.BackColor = $colors.StatusBar
    $statusLabel.BackColor = $colors.StatusBar
    $statusLabel.ForeColor = [System.Drawing.Color]::White
    $versionLabel.BackColor = $colors.StatusBar
    $versionLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)
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
    $btnSmartUpdate.Text = Get-LocalizedString "BtnSmartUpdate"
    $btnScan.Text = Get-LocalizedString "BtnScan"
    $btnBackup.Text = Get-LocalizedString "BtnBackup"
    $btnInstall.Text = Get-LocalizedString "BtnInstall"
    $btnCancel.Text = Get-LocalizedString "BtnCancel"
    $menuOpenLogs.Text = Get-LocalizedString "BtnOpenLogs"
    $menuSmartUpdate.Text = Get-LocalizedString "BtnSmartUpdate"
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
    $historyList.FullRowSelect = true
    $historyList.GridLines = false
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
    $schedForm.Controls.Add($contentPanel)
    $schedForm.Controls.Add($header)
    $lblFreq = New-Object Windows.Forms.Label
    $lblFreq.Text = "Frequency:"
    $lblFreq.Location = '0,10'
    $lblFreq.Size = '100,25'
    $lblFreq.ForeColor = $colors.Text
    $contentPanel.Controls.Add($lblFreq)
    $cmbFreq = New-Object Windows.Forms.ComboBox
    $cmbFreq.Location = '170,8'
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
    $txtTime.Location = '170,48'
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
        Add-StatusUI $form $status "Filter applied: Class='$($txtClass.Text)', Manufacturer='$($txtMfr.Text)'"
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
        Add-StatusUI $form $status "Filters cleared."
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
    $settingsForm.Size = '420,240'
    $settingsForm.StartPosition = "CenterParent"
    $settingsForm.BackColor = $colors.Background
    $settingsForm.FormBorderStyle = 'FixedDialog'
    $settingsForm.MaximizeBox = false
    $settingsForm.TopMost = true
    $settingsForm.Font = New-Object Drawing.Font("Segoe UI", 9.5)
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
    $settingsForm.Controls.Add($contentPanel)
    $settingsForm.Controls.Add($header)
    # Removed proxy UI elements here
    $btnPanel = New-Object Windows.Forms.Panel
    $btnPanel.Location = '0,50'
    $btnPanel.Size = '360,40'
    $contentPanel.Controls.Add($btnPanel)
    $btnClose = New-Object Windows.Forms.Button
    $btnClose.Text = Get-LocalizedString "Close"
    $btnClose.Location = '230,0'
    $btnClose.Size = '100,36'
    $btnClose.FlatStyle = 'Flat'
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.BackColor = $colors.Secondary
    $btnClose.ForeColor = $colors.Text
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $settingsForm.Close() })
    $btnPanel.Controls.Add($btnClose)
    $settingsForm.ShowDialog($form) | Out-Null
}

# -------------------------
# Event Handlers
# -------------------------
$btnSmartUpdate.Add_Click({ Start-BackgroundTask -Name "SmartUpdate" })
$menuSmartUpdate.Add_Click({ $btnSmartUpdate.PerformClick() })

$btnScan.Add_Click({ Start-BackgroundTask -Name "ScanDrivers" })
$menuScan.Add_Click({ $btnScan.PerformClick() })

$btnBackup.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = Get-LocalizedString "SelectBackupFolder"
    if ($fbd.ShowDialog($form) -eq "OK") {
        Start-BackgroundTask -Name "BackupDrivers" -TaskArgs @($fbd.SelectedPath)
    }
})
$menuBackup.Add_Click({ $btnBackup.PerformClick() })

$btnInstall.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = Get-LocalizedString "SelectDriverFolder"
    if ($fbd.ShowDialog($form) -eq "OK") {
        Start-BackgroundTask -Name "InstallDrivers" -TaskArgs @($fbd.SelectedPath)
    }
})
$menuInstall.Add_Click({ $btnInstall.PerformClick() })

$btnCancel.Add_Click({
    if ($null -ne $global:CurrentJob) {
        Stop-Job -Id $global:CurrentJob.Id -PassThru | Remove-Job -Force -ErrorAction SilentlyContinue
        $global:CurrentJob = $null
        $timer.Stop()
        Add-StatusUI $form $status "`r`n$(Get-LocalizedString 'TaskCancelled')"
        $statusLabel.Text = "  Task Cancelled"
    } else {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "NoTaskToCancel"), "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$menuCancel.Add_Click({ $btnCancel.PerformClick() })

$menuOpenLogs.Add_Click({
    if (Test-Path $LogBase) {
        [System.Diagnostics.Process]::Start($LogBase)
    } else {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "LogFolderMissing"), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$menuExit.Add_Click({ $form.Close() })
$menuRestorePoint.Add_Click({
    if (New-RestorePoint) {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "RestorePointCreated"), "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show((Get-LocalizedString "RestorePointFailed"), "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$menuSchedule.Add_Click({ Show-ScheduleDialog })
$menuFilters.Add_Click({ Show-FiltersDialog })
$menuHistory.Add_Click({ Show-HistoryDialog })
$menuSettingsTop.Add_Click({ Show-SettingsDialog }) # This will now show a simplified settings dialog
$menuToggleTheme.Add_Click({
    Set-Theme -Dark (-not $global:DarkModeEnabled)
})

foreach ($item in $menuLanguage.DropDownItems) {
    $item.Add_Click({
        $global:CurrentLanguage = $item.Tag
        Export-Settings
        Add-StatusUI $form $status "Language changed to $($item.Tag)... Restart app to apply fully."
        Update-UILanguage
    })
}

$form.Add_Resize({
    Update-ButtonContainerPadding
})
$form.Add_Shown({
    Update-ButtonContainerPadding
})

$form.Add_Closing({
    if ($null -ne $global:CurrentJob) {
        $btnCancel.PerformClick() # Attempt to cancel running job
    }
    Export-Settings # Save settings on exit
})

Set-Theme -Dark $global:DarkModeEnabled
$form.ShowDialog() | Out-Null

if ($null -ne $global:CurrentJob) {
    Remove-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
}
exit 0