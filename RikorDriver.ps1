# Driver Updater Rikor - Minimal Scan & Install Module
# Run as Administrator
param(
    [switch]$Silent,
    [string]$Task = "",
    [string]$Language = "ru",
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
        BtnScan = "Scan for Missing Drivers"
        BtnInstall = "Install From Folder"
        PermissionError = "Please run this script as Administrator."
        StartingTask = "-> Starting task:"
        TaskFinished = "=== Task finished:"
        SelectDriverFolder = "Select folder containing driver .inf files"
        PSWU_Ensuring = "Ensuring PSWindowsUpdate module is available..."
        PSWU_Downloading = "Downloading PSWindowsUpdate module from PowerShell Gallery..."
        PSWU_Extracting = "Extracting PSWindowsUpdate module..."
        PSWU_Importing = "Importing PSWindowsUpdate module..."
        PSWU_Available = "PSWindowsUpdate module is available."
        PSWU_Failed = "Failed to ensure PSWindowsUpdate module availability."
        Scan_MissingDriversFound = "Found {0} missing driver(s):"
        Scan_NoMissingDrivers = "No missing driver updates found."
        Scan_ExportCSV = "Exporting missing drivers list to CSV..."
        Scan_Failed = "Failed to scan for missing drivers."
    }
    "ru" = @{
        AppTitle = "Установщик драйверов Rikor"
        BtnScan = "Сканировать недостающие драйверы"
        BtnInstall = "Установить из папки"
        PermissionError = "Запустите этот скрипт от имени администратора."
        StartingTask = "-> Запуск задачи:"
        TaskFinished = "=== Задача завершена:"
        SelectDriverFolder = "Выберите папку с файлами драйверов (.inf)"
        PSWU_Ensuring = "Проверка наличия модуля PSWindowsUpdate..."
        PSWU_Downloading = "Загрузка модуля PSWindowsUpdate с PowerShell Gallery..."
        PSWU_Extracting = "Распаковка модуля PSWindowsUpdate..."
        PSWU_Importing = "Импорт модуля PSWindowsUpdate..."
        PSWU_Available = "Модуль PSWindowsUpdate доступен."
        PSWU_Failed = "Не удалось обеспечить доступность модуля PSWindowsUpdate."
        Scan_MissingDriversFound = "Найдено {0} недостающих драйверов:"
        Scan_NoMissingDrivers = "Недостающие обновления драйверов не найдены."
        Scan_ExportCSV = "Экспорт списка недостающих драйверов в CSV..."
        Scan_Failed = "Не удалось отсканировать недостающие драйверы."
    }
}

function Get-LocalizedString([string]$key, [array]$args = $null) {
    $lang = $global:CurrentLanguage
    if (-not $global:Languages.ContainsKey($lang)) { $lang = "en" }

    $string = if ($global:Languages[$lang].ContainsKey($key)) {
        $global:Languages[$lang][$key]
    } else {
        $global:Languages["en"][$key]
    }

    if ($args) {
        return $string -f $args
    }
    return $string
}

$global:CurrentLanguage = $Language

# -------------------------
# Require Admin Privilege
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
# Fix for PSScriptRoot being empty
# -------------------------
if (-not $PSScriptRoot) {
    # Try to get the script path from the call stack
    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
    if (-not $PSScriptRoot) {
        # Fallback: use current directory
        $PSScriptRoot = $PWD.Path
    }
}

# -------------------------
# Ensure PSWindowsUpdate Module
# -------------------------
function Ensure-PSWindowsUpdateModule {
    param($statusRef)

    $moduleName = "PSWindowsUpdate"
    $modulePath = Join-Path $PSScriptRoot "Modules\$moduleName"

    # Check if already imported
    if (Get-Module -Name $moduleName -ErrorAction SilentlyContinue) {
        Add-StatusUI $statusRef "[INFO] $moduleName module is already loaded."
        return $true
    }

    # Check local path
    if (Test-Path "$modulePath\$moduleName.psd1") {
        Add-StatusUI $statusRef "[INFO] $moduleName module found locally. Importing..."
        try {
            Import-Module $modulePath -Force -ErrorAction Stop
            Add-StatusUI $statusRef (Get-LocalizedString "PSWU_Available")
            return $true
        } catch {
            Add-StatusUI $statusRef "[ERROR] Failed to import $moduleName from local path: $_"
        }
    }

    Add-StatusUI $statusRef (Get-LocalizedString "PSWU_Ensuring")

    try {
        # Download and extract module
        $moduleZipUrl = "https://www.powershellgallery.com/api/v2/package/PSWindowsUpdate"
        $moduleZipPath = Join-Path $env:TEMP "$moduleName.zip"
        $localModuleDir = Join-Path $PSScriptRoot "Modules"

        if (-not (Test-Path $localModuleDir)) { New-Item -ItemType Directory -Path $localModuleDir -Force | Out-Null }

        Add-StatusUI $statusRef (Get-LocalizedString "PSWU_Downloading")
        Invoke-WebRequest -Uri $moduleZipUrl -OutFile $moduleZipPath -UseBasicParsing -ErrorAction Stop

        Add-StatusUI $statusRef (Get-LocalizedString "PSWU_Extracting")
        Expand-Archive -Path $moduleZipPath -DestinationPath "$localModuleDir\_temp" -Force -ErrorAction Stop

        $extractedContentPath = Get-ChildItem -Path "$localModuleDir\_temp" -Directory -Filter "$moduleName*" | Select-Object -First 1
        if ($extractedContentPath) {
            $versionFolder = Get-ChildItem -Path $extractedContentPath.FullName -Directory | Sort-Object Name -Descending | Select-Object -First 1
            if ($versionFolder) {
                Move-Item -Path "$versionFolder.FullName\*" -Destination $modulePath -Force
            } else {
                Move-Item -Path "$extractedContentPath.FullName\*" -Destination $modulePath -Force
            }
        }

        Remove-Item "$localModuleDir\_temp" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item $moduleZipPath -Force -ErrorAction SilentlyContinue

        Add-StatusUI $statusRef (Get-LocalizedString "PSWU_Importing")
        Import-Module $modulePath -Force -ErrorAction Stop
        Add-StatusUI $statusRef (Get-LocalizedString "PSWU_Available")
        return $true
    } catch {
        Add-StatusUI $statusRef "[ERROR] $(Get-LocalizedString 'PSWU_Failed'): $_"
        return $false
    }
}

# -------------------------
# Helper Function for Console Output
# -------------------------
function Add-StatusUI {
    param($statusControl, $text)
    if ($statusControl) {
        $method = [System.Windows.Forms.MethodInvoker]{
            $statusControl.AppendText("$text`r`n")
            $statusControl.ScrollToCaret()
        }
        if ($statusControl.InvokeRequired) {
            $statusControl.Invoke($method)
        } else {
            $method.Invoke()
        }
    }
}

# -------------------------
# Scan for Missing Drivers
# -------------------------
function Start-ScanDrivers {
    param($statusControl)
    Add-StatusUI $statusControl "$(Get-LocalizedString 'StartingTask') ScanDrivers"

    if (-not (Ensure-PSWindowsUpdateModule -statusRef $statusControl)) {
        Add-StatusUI $statusControl "[ERROR] Failed to load PSWindowsUpdate module."
        return
    }

    try {
        Add-StatusUI $statusControl (Get-LocalizedString "WU_Probing")
        $missingDrivers = Get-WindowsUpdate -Driver -NotInstalled -ErrorAction Stop

        $filteredDrivers = $missingDrivers
        if ($FilterClass) {
            Add-StatusUI $statusControl "Applying class filter: $FilterClass"
            $filteredDrivers = $filteredDrivers | Where-Object { $_.Categories -join ', ' -like "*$FilterClass*" }
        }
        if ($FilterManufacturer) {
            Add-StatusUI $statusControl "Applying manufacturer filter: $FilterManufacturer"
            $filteredDrivers = $filteredDrivers | Where-Object { $_.Manufacturer -like "*$FilterManufacturer*" }
        }

        if ($filteredDrivers.Count -eq 0) {
            Add-StatusUI $statusControl (Get-LocalizedString "Scan_NoMissingDrivers")
        } else {
            Add-StatusUI $statusControl (Get-LocalizedString "Scan_MissingDriversFound", @($filteredDrivers.Count))
            foreach ($driver in $filteredDrivers) {
                Add-StatusUI $statusControl "  - Title: $($driver.Title) (Manufacturer: $($driver.Manufacturer), Version: $($driver.Version))"
            }

            Add-StatusUI $statusControl (Get-LocalizedString "Scan_ExportCSV")
            $csvPath = Join-Path $env:TEMP "MissingDrivers_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
            $filteredDrivers | Select-Object Title, Manufacturer, @{Name='Categories';Expression={ $_.Categories -join ', ' }}, Version, LastDeploymentChangeTime | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Add-StatusUI $statusControl "Exported to: $csvPath"
        }
    } catch {
        Add-StatusUI $statusControl "[ERROR] $(Get-LocalizedString 'Scan_Failed'): $_"
    } finally {
        Add-StatusUI $statusControl "`r`n$(Get-LocalizedString 'TaskFinished') ScanDrivers"
    }
}

# -------------------------
# Install Drivers from Folder
# -------------------------
function Start-InstallDrivers {
    param($statusControl)
    Add-StatusUI $statusControl "$(Get-LocalizedString 'StartingTask') InstallDrivers"

    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = Get-LocalizedString "SelectDriverFolder"
    if ($fbd.ShowDialog() -ne "OK") {
        Add-StatusUI $statusControl "[INFO] Installation canceled by user."
        return
    }

    $folder = $fbd.SelectedPath
    Add-StatusUI $statusControl "Installing drivers from: $folder"

    try {
        if (-not (Test-Path $folder)) {
            Add-StatusUI $statusControl "[ERROR] Folder not found: $folder"
            return
        }

        $infFiles = Get-ChildItem -Path $folder -Recurse -Include "*.inf"
        if ($infFiles.Count -eq 0) {
            Add-StatusUI $statusControl "[ERROR] No .inf driver files found in folder"
            return
        }

        Add-StatusUI $statusControl "Found $($infFiles.Count) .inf files. Installing..."

        $successCount = 0; $failCount = 0
        foreach ($inf in $infFiles) {
            Add-StatusUI $statusControl "Installing $($inf.Name)..."
            try {
                & pnputil.exe /add-driver $inf.FullName /install /force 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    $successCount++
                } else {
                    $failCount++
                    Add-StatusUI $statusControl " -> Failed (Code: $LASTEXITCODE)"
                }
            } catch {
                $failCount++
                Add-StatusUI $statusControl " -> Failed (Exception: $_)"
            }
        }

        Add-StatusUI $statusControl "Installation complete: $successCount successful, $failCount failed."
    } catch {
        Add-StatusUI $statusControl "[ERROR] Installation failed: $_"
    } finally {
        Add-StatusUI $statusControl "`r`n$(Get-LocalizedString 'TaskFinished') InstallDrivers"
    }
}

# -------------------------
# Silent Mode for ScanDrivers
# -------------------------
if ($Silent -and $Task -eq "ScanDrivers") {
    $logFile = Join-Path $env:TEMP "ScanDrivers_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    function Write-SilentLog($msg) {
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Add-Content -Path $logFile -Value "$timestamp - $msg"
    }

    Write-SilentLog "Starting silent mode: ScanDrivers"
    Write-SilentLog (Get-LocalizedString "PSWU_Ensuring")

    if (-not (Ensure-PSWindowsUpdateModule -statusRef $null)) {
        Write-SilentLog "[ERROR] Failed to load PSWindowsUpdate module."
        exit 1
    }

    try {
        Write-SilentLog (Get-LocalizedString "WU_Probing")
        $missingDrivers = Get-WindowsUpdate -Driver -NotInstalled -ErrorAction Stop

        $filteredDrivers = $missingDrivers
        if ($FilterClass) {
            Write-SilentLog "Applying class filter: $FilterClass"
            $filteredDrivers = $filteredDrivers | Where-Object { $_.Categories -join ', ' -like "*$FilterClass*" }
        }
        if ($FilterManufacturer) {
            Write-SilentLog "Applying manufacturer filter: $FilterManufacturer"
            $filteredDrivers = $filteredDrivers | Where-Object { $_.Manufacturer -like "*$FilterManufacturer*" }
        }

        if ($filteredDrivers.Count -eq 0) {
            Write-SilentLog (Get-LocalizedString "Scan_NoMissingDrivers")
        } else {
            Write-SilentLog (Get-LocalizedString "Scan_MissingDriversFound", @($filteredDrivers.Count))
            foreach ($driver in $filteredDrivers) {
                Write-SilentLog "  - Title: $($driver.Title) (Manufacturer: $($driver.Manufacturer), Version: $($driver.Version))"
            }

            Write-SilentLog (Get-LocalizedString "Scan_ExportCSV")
            $csvPath = Join-Path $env:TEMP "MissingDrivers_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
            $filteredDrivers | Select-Object Title, Manufacturer, @{Name='Categories';Expression={ $_.Categories -join ', ' }}, Version, LastDeploymentChangeTime | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-SilentLog "Exported to: $csvPath"
        }
    } catch {
        Write-SilentLog "[ERROR] $(Get-LocalizedString 'Scan_Failed'): $_"
    }

    Write-SilentLog "Silent mode completed."
    exit 0
}

# -------------------------
# GUI Form (Minimal)
# -------------------------
$form = New-Object Windows.Forms.Form
$form.Text = Get-LocalizedString "AppTitle"
$form.Size = '800,600'
$form.StartPosition = "CenterScreen"
$form.Font = New-Object Drawing.Font("Segoe UI", 9.5)

# Menu Strip
$menuStrip = New-Object Windows.Forms.MenuStrip
$menuFile = New-Object Windows.Forms.ToolStripMenuItem("&File")
$menuExit = New-Object Windows.Forms.ToolStripMenuItem("Exit")
$menuExit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$menuFile.DropDownItems.AddRange(@($menuExit))
$menuActions = New-Object Windows.Forms.ToolStripMenuItem("&Actions")
$menuScan = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnScan"))
$menuInstall = New-Object Windows.Forms.ToolStripMenuItem((Get-LocalizedString "BtnInstall"))
$menuActions.DropDownItems.AddRange(@($menuScan, $menuInstall))
$menuStrip.Items.AddRange(@($menuFile, $menuActions))
$form.MainMenuStrip = $menuStrip

# Toolbar Panel
$toolbarPanel = New-Object Windows.Forms.Panel
$toolbarPanel.Dock = 'Top'
$toolbarPanel.Height = 60
$buttonContainer = New-Object Windows.Forms.FlowLayoutPanel
$buttonContainer.Dock = 'Fill'
$buttonContainer.FlowDirection = 'LeftToRight'
$buttonContainer.WrapContents = $false
$buttonContainer.AutoSize = $false
$buttonContainer.Padding = '10,10,10,10'
$toolbarPanel.Controls.Add($buttonContainer)

# Buttons
$btnScan = New-Object Windows.Forms.Button
$btnScan.Text = Get-LocalizedString "BtnScan"
$btnScan.Width = 200
$btnScan.Height = 38
$btnScan.FlatStyle = "Flat"
$btnScan.FlatAppearance.BorderSize = 0
$btnScan.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnScan.Font = New-Object Drawing.Font("Segoe UI Semibold", 9)
$btnScan.Add_Click({ Start-ScanDrivers -statusControl $status })
$buttonContainer.Controls.Add($btnScan)

$btnInstall = New-Object Windows.Forms.Button
$btnInstall.Text = Get-LocalizedString "BtnInstall"
$btnInstall.Width = 200
$btnInstall.Height = 38
$btnInstall.FlatStyle = "Flat"
$btnInstall.FlatAppearance.BorderSize = 0
$btnInstall.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnInstall.Font = New-Object Drawing.Font("Segoe UI Semibold", 9)
$btnInstall.Add_Click({ Start-InstallDrivers -statusControl $status })
$buttonContainer.Controls.Add($btnInstall)

# Status Bar
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

# Output Console
$contentPanel = New-Object Windows.Forms.Panel
$contentPanel.Dock = 'Fill'
$contentPanel.Padding = '10,10,10,10'
$form.Controls.Add($contentPanel)
$form.Controls.Add($statusBar)
$form.Controls.Add($toolbarPanel)
$form.Controls.Add($menuStrip)

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

# Event Handlers
$menuExit.Add_Click({ $form.Close() })

# Show Form
$form.ShowDialog() | Out-Null
exit 0