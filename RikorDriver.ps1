# Driver Updater Rikor v 1.0
# Run as Administrator
# Usage: irm https://raw.githubusercontent.com/h4rd1ns0mn14/RikorDriver/refs/heads/main/RikorDriver.ps1 | iex
# Silent mode: .\RikorDriver.ps1 -Silent -Task "CheckDriverUpdates"

param(
    [switch]$Silent,
    [string]$Task = "",
    [string]$ProxyAddress = "",
    [string]$FilterClass = "",
    [string]$FilterManufacturer = ""
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -------------------------
# Require Admin
# -------------------------
function Assert-AdminPrivilege {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show("Please run this script as Administrator.", "Permission Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        exit 1
    }
}
Assert-AdminPrivilege

# -------------------------
# Globals and paths
# -------------------------
$AppTitle = "Rikor Driver Installer"
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
    param([string]$Description = "Rikor Driver Installer Restore Point")
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
$ScheduledTaskName = "RikorDriverInstaller_ScheduledCheck"

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
        
        # Determine script path. If running via IEX, we need to save the script to disk first.
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { 
            # We are running from memory (IEX). Save to Documents folder to persist it.
            $scriptPath = Join-Path $LogBase "RikorDriver.ps1"
            
            # Since 'irm | iex' is the usage, let's assume we download the latest version to install it.
            $downloadUrl = "https://raw.githubusercontent.com/h4rd1ns0mn14/RikorDriver/refs/heads/main/RikorDriver.ps1"
            try {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $scriptPath -UseBasicParsing
            } catch {
                throw "Could not save script to $scriptPath for scheduling. Please download the script manually."
            }
        }
        
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -Silent -Task DownloadAndInstallDrivers"
        
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
        "DownloadAndInstallDrivers" {
            Write-SilentLog "Silent mode: Downloading and installing drivers from Rikor archive..."
            
            # Get computer model
            $computerModel = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
            Write-SilentLog "Detected computer model: $computerModel"
            
            # Load Nextcloud URLs from online JSON file
            $modelsFileUrl = "https://nc.rikor.com/index.php/s/BfBKYyW9HdoFfz9/download"
            $nextcloudUrls = @{}
            
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                
                $modelsJsonString = $webClient.DownloadString($modelsFileUrl)
                $tempObj = $modelsJsonString | ConvertFrom-Json
                $nextcloudUrls = @{}
                foreach($key in $tempObj.PSObject.Properties.Name) { 
                    $val = $tempObj.$key
                    if ($val -is [string]) {
                        $nextcloudUrls[$key] = $val
                    } elseif ($val.url) {
                        $nextcloudUrls[$key] = $val.url
                    }
                }
                
                Write-SilentLog "Loaded models mapping from online file"
            } catch {
                Write-SilentLog "[WARNING] Failed to load online models.json: $_"
                # Local fallback is only possible if script exists on disk with models.json
                if ($PSScriptRoot) {
                    $modelsFilePath = Join-Path $PSScriptRoot "models.json"
                    if (Test-Path $modelsFilePath) {
                        try {
                            $tempObj = Get-Content -Path $modelsFilePath -Raw | ConvertFrom-Json
                            $nextcloudUrls = @{}
                            foreach($key in $tempObj.PSObject.Properties.Name) { 
                                $val = $tempObj.$key
                                if ($val -is [string]) {
                                    $nextcloudUrls[$key] = $val
                                } elseif ($val.url) {
                                    $nextcloudUrls[$key] = $val.url
                                }
                            }
                            Write-SilentLog "Loaded models mapping from local fallback file"
                        } catch {
                             Write-SilentLog "[ERROR] Local models.json invalid: $_"
                        }
                    }
                }
            }
            
            # Determine the appropriate URL based on model
            $zipUrl = $null
            $rikorServerAvailable = $false
            
            if ($nextcloudUrls.ContainsKey($computerModel)) {
                $zipUrl = $nextcloudUrls[$computerModel]
                Write-SilentLog "Using Rikor Server for download"
                
                try {
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    $request = [System.Net.WebRequest]::Create($zipUrl)
                    $request.Method = "HEAD"
                    $request.Timeout = 15000 
                    $request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                    
                    $response = $request.GetResponse()
                    if ($response.StatusCode -eq 200) {
                        $rikorServerAvailable = $true
                    }
                    $response.Close()
                } catch {
                    Write-SilentLog "[INFO] Rikor server is not accessible: $_"
                }
            } else {
                Write-SilentLog "Model '$computerModel' not in predefined list, checking Microsoft Update"
            }

            if (-not $rikorServerAvailable) {
                # Fallback to Microsoft Update
                Write-SilentLog "[INFO] Rikor server is not available or model unknown. Checking Microsoft Update..."
                
                try {
                    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
                    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
                    $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver'")
                    
                    if ($SearchResult.Updates.Count -eq 0) {
                        Write-SilentLog "No driver updates available from Microsoft Update"
                    } else {
                        Write-SilentLog "Found $($SearchResult.Updates.Count) driver update(s) available."
                        
                        $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
                        foreach ($Update in $SearchResult.Updates) {
                            if ($Update.EulaAccepted -eq $false) { $Update.AcceptEula() }
                            $UpdatesToDownload.Add($Update) | Out-Null
                        }
                        
                        Write-SilentLog "Downloading updates..."
                        $Downloader = $UpdateSession.CreateUpdateDownloader()
                        $Downloader.Updates = $UpdatesToDownload
                        $Downloader.Download()
                        
                        Write-SilentLog "Installing updates..."
                        $Installer = $UpdateSession.CreateUpdateInstaller()
                        
                        # Filter for downloaded updates
                        $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                        foreach ($Update in $SearchResult.Updates) {
                            if ($Update.IsDownloaded) { $UpdatesToInstall.Add($Update) | Out-Null }
                        }
                        
                        if ($UpdatesToInstall.Count -gt 0) {
                            $Installer.Updates = $UpdatesToInstall
                            $InstallResult = $Installer.Install()
                            Write-SilentLog "Installation completed. ResultCode: $($InstallResult.ResultCode). RebootRequired: $($InstallResult.RebootRequired)"
                        }
                    }
                } catch {
                    Write-SilentLog "[ERROR] Microsoft Update failed: $_"
                }
                
                Write-SilentLog "Completed"
                return
            }

            # Download and Install from Rikor
            $tempDir = Join-Path $env:TEMP "RikorDriversTempSilent_$(Get-Date -Format 'yyyyMMddHHmmss')"
            try {
                New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                $zipPath = Join-Path $tempDir "drivers.zip"
                
                Write-SilentLog "Downloading drivers archive..."
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                $webClient.DownloadFile($zipUrl, $zipPath)
                
                if (-not (Test-Path $zipPath) -or (Get-Item $zipPath).Length -eq 0) {
                    throw "Downloaded ZIP file is missing or empty."
                }
                
                $extractDir = Join-Path $tempDir "ExtractedDrivers"
                Write-SilentLog "Extracting to: $extractDir"
                Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
                
                $infFiles = Get-ChildItem -Path $extractDir -Recurse -Include "*.inf" -ErrorAction SilentlyContinue
                Write-SilentLog "Found $($infFiles.Count) .inf file(s)."
                
                $successCount = 0
                $failCount = 0
                
                foreach ($inf in $infFiles) {
                    Write-SilentLog "Installing: $($inf.Name)"
                    $out = & pnputil.exe /add-driver $inf.FullName /install /force 2>&1
                    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 259) { # 259 = No more items (up to date)
                        $successCount++
                    } else {
                        $failCount++
                        Write-SilentLog "Failed: $out"
                    }
                }
                Write-SilentLog "Result: $successCount installed/up-to-date, $failCount failed."
                
            } catch {
                Write-SilentLog "[ERROR] $_"
            } finally {
                if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
            }
            Write-SilentLog "Completed"
        }
        
        default {
            Write-SilentLog "Unknown task: $Task"
        }
    }
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
        Surface = [System.Drawing.Color]::FromArgb(32, 32, 32)
        SurfaceHover = [System.Drawing.Color]::FromArgb(48, 48, 48)
        Primary = [System.Drawing.Color]::FromArgb(77, 73, 190) # Pantone 2368 C
        PrimaryHover = [System.Drawing.Color]::FromArgb(97, 93, 210)
        Secondary = [System.Drawing.Color]::FromArgb(50, 50, 50)
        Text = [System.Drawing.Color]::FromArgb(245, 245, 245)
        TextSecondary = [System.Drawing.Color]::FromArgb(160, 160, 160)
        Border = [System.Drawing.Color]::FromArgb(40, 40, 40)
        Success = [System.Drawing.Color]::FromArgb(76, 175, 80)
        Warning = [System.Drawing.Color]::FromArgb(255, 152, 0)
        Error = [System.Drawing.Color]::FromArgb(244, 67, 54)
        MenuBar = [System.Drawing.Color]::FromArgb(24, 24, 24)
        StatusBar = [System.Drawing.Color]::FromArgb(77, 73, 190) # Pantone 2368 C
    }
    # Light Theme
    Light = @{
        Background = [System.Drawing.Color]::FromArgb(250, 250, 252) # Cooler white
        Surface = [System.Drawing.Color]::FromArgb(255, 255, 255)
        SurfaceHover = [System.Drawing.Color]::FromArgb(240, 242, 245)
        Primary = [System.Drawing.Color]::FromArgb(77, 73, 190) # Pantone 2368 C
        PrimaryHover = [System.Drawing.Color]::FromArgb(97, 93, 210)
        Secondary = [System.Drawing.Color]::FromArgb(235, 235, 240)
        Text = [System.Drawing.Color]::FromArgb(30, 30, 35)
        TextSecondary = [System.Drawing.Color]::FromArgb(100, 100, 110)
        Border = [System.Drawing.Color]::FromArgb(230, 230, 235)
        Success = [System.Drawing.Color]::FromArgb(67, 160, 71)
        Warning = [System.Drawing.Color]::FromArgb(251, 140, 0)
        Error = [System.Drawing.Color]::FromArgb(229, 57, 53)
        MenuBar = [System.Drawing.Color]::FromArgb(255, 255, 255)
        StatusBar = [System.Drawing.Color]::FromArgb(77, 73, 190) # Pantone 2368 C
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
$form.Size = '1100,750'
$form.MinimumSize = '1050,720'
$form.StartPosition = "CenterScreen"
$form.Font = New-Object Drawing.Font("Segoe UI", 10)

# -------------------------
# Menu Strip
# -------------------------
$menuStrip = New-Object Windows.Forms.MenuStrip
$menuStrip.Padding = '6,2,0,2'

# File Menu
$menuFile = New-Object Windows.Forms.ToolStripMenuItem
$menuFile.Text = "&File"
$menuOpenLogs = New-Object Windows.Forms.ToolStripMenuItem
$menuOpenLogs.Text = "Open Logs"
$menuOpenLogs.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::L
$menuSeparator1 = New-Object Windows.Forms.ToolStripSeparator
$menuExit = New-Object Windows.Forms.ToolStripMenuItem
$menuExit.Text = "Exit"
$menuExit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$menuFile.DropDownItems.AddRange(@($menuOpenLogs, $menuSeparator1, $menuExit))

# Actions Menu
$menuActions = New-Object Windows.Forms.ToolStripMenuItem
$menuActions.Text = "&Actions"

$menuDownloadAndInstall = New-Object Windows.Forms.ToolStripMenuItem
$menuDownloadAndInstall.Text = "Download and Install Rikor drivers"
$menuDownloadAndInstall.ShortcutKeys = [System.Windows.Forms.Keys]::F6

$menuSeparator2 = New-Object Windows.Forms.ToolStripSeparator

$menuBackup = New-Object Windows.Forms.ToolStripMenuItem
$menuBackup.Text = "Backup Drivers"
$menuBackup.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::B

$menuInstall = New-Object Windows.Forms.ToolStripMenuItem
$menuInstall.Text = "Install From Folder"
$menuInstall.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::I

$menuSeparator3 = New-Object Windows.Forms.ToolStripSeparator
$menuCancel = New-Object Windows.Forms.ToolStripMenuItem
$menuCancel.Text = "Cancel Task"
$menuCancel.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Q

$menuActions.DropDownItems.AddRange(@($menuDownloadAndInstall, $menuSeparator2, $menuBackup, $menuInstall, $menuSeparator3, $menuCancel))

# Tools Menu
$menuTools = New-Object Windows.Forms.ToolStripMenuItem
$menuTools.Text = "&Tools"
$menuRestorePoint = New-Object Windows.Forms.ToolStripMenuItem
$menuRestorePoint.Text = "Create Restore Point"
$menuSchedule = New-Object Windows.Forms.ToolStripMenuItem
$menuSchedule.Text = "Schedule Updates"
$menuFilters = New-Object Windows.Forms.ToolStripMenuItem
$menuFilters.Text = "Filters"
$menuSeparator4 = New-Object Windows.Forms.ToolStripSeparator
$menuHistory = New-Object Windows.Forms.ToolStripMenuItem
$menuHistory.Text = "Update History"
$menuHistory.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::H
$menuTools.DropDownItems.AddRange(@($menuRestorePoint, $menuSchedule, $menuFilters, $menuSeparator4, $menuHistory))

# View Menu
$menuView = New-Object Windows.Forms.ToolStripMenuItem
$menuView.Text = "&View"
$menuToggleTheme = New-Object Windows.Forms.ToolStripMenuItem
$menuToggleTheme.Text = "Dark Mode"
$menuToggleTheme.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::T
$menuView.DropDownItems.AddRange(@($menuToggleTheme))

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

$buttonContainer = New-Object Windows.Forms.FlowLayoutPanel
$buttonContainer.Dock = 'Fill'
$buttonContainer.FlowDirection = 'LeftToRight'
$buttonContainer.WrapContents = $false
$buttonContainer.AutoSize = $false
$buttonContainer.Padding = '0,8,0,8'
$toolbarPanel.Controls.Add($buttonContainer)

function Update-ButtonContainerPadding {
    $totalButtonWidth = 240 + 120 + 140 + 110 + (3 * 12)
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
    $btn.Height = 42
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Font = New-Object Drawing.Font("Segoe UI Semibold", 10)
    $btn.Tag = $Primary
    $btn.TextAlign = 'MiddleCenter'
    $btn.Region = New-RoundedRegion -Width $Width -Height 42 -Radius 12
    return $btn
}

$btnDownloadAndInstall = New-ModernButton -Text "Download and Install Rikor drivers" -Width 240 -Primary $true
$btnDownloadAndInstall.Margin = '0,0,12,0'

$btnBackup = New-ModernButton -Text "Backup" -Width 120
$btnBackup.Margin = '0,0,12,0'

$btnInstall = New-ModernButton -Text "Install From Disk" -Width 140
$btnInstall.Margin = '0,0,12,0'

$btnCancel = New-ModernButton -Text "Cancel" -Width 110
$btnCancel.Margin = '0,0,0,0'

$buttonContainer.Controls.AddRange(@($btnDownloadAndInstall, $btnBackup, $btnInstall, $btnCancel))

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
$versionLabel.Text = "v3.0  "
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
$contentPanel.Padding = '24,24,24,12'

$form.Controls.Add($contentPanel)
$form.Controls.Add($statusBar)
$form.Controls.Add($toolbarSeparator)
$form.Controls.Add($toolbarPanel)
$form.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip

$statusBorderPanel = New-Object Windows.Forms.Panel
$statusBorderPanel.Dock = 'Fill'
$statusBorderPanel.Padding = '1,1,1,1' # Minimal border
$contentPanel.Controls.Add($statusBorderPanel)

$status = New-Object Windows.Forms.RichTextBox
$status.Multiline = $true
$status.ReadOnly = $true
$status.Dock = 'Fill'
$status.ScrollBars = 'Vertical'
$status.BorderStyle = 'None'
$status.Font = New-Object Drawing.Font("Consolas", 10)
$statusBorderPanel.Controls.Add($status)

# Progress Panel (Overall Only)
$progressPanel = New-Object Windows.Forms.Panel
$progressPanel.Dock = 'Bottom'
$progressPanel.Height = 36
$progressPanel.Padding = '0,8,0,0'
$contentPanel.Controls.Add($progressPanel)

# Overall Progress
$overallProgressPanel = New-Object Windows.Forms.Panel
$overallProgressPanel.Dock = 'Fill'
$overallProgressPanel.Padding = '0,0,0,4'
$progressPanel.Controls.Add($overallProgressPanel)

$overallProgressBorder = New-Object Windows.Forms.Panel
$overallProgressBorder.Dock = 'Fill'
$overallProgressBorder.Padding = '1,1,1,1'
$overallProgressPanel.Controls.Add($overallProgressBorder)

$progress = New-Object Windows.Forms.ProgressBar
$progress.Dock = 'Fill'
$progress.Style = 'Continuous'
$progress.Value = 0
$overallProgressBorder.Controls.Add($progress)

$overallLabel = New-Object Windows.Forms.Label
$overallLabel.Text = "Overall: 0%"
$overallLabel.Dock = 'Right'
$overallLabel.Width = 120
$overallLabel.TextAlign = 'MiddleRight'
$overallLabel.Font = New-Object Drawing.Font("Segoe UI", 8.5)
$overallProgressPanel.Controls.Add($overallLabel)

$headerLabel = New-Object Windows.Forms.Label
$headerLabel.Text = "Output Console"
$headerLabel.Dock = 'Top'
$headerLabel.Height = 26
$headerLabel.Font = New-Object Drawing.Font("Segoe UI Semibold", 9.5)
$headerLabel.TextAlign = 'MiddleLeft'
$contentPanel.Controls.Add($headerLabel)

$timer = New-Object Windows.Forms.Timer
$timer.Interval = 500

# -------------------------
# Background job helpers
# -------------------------
function Start-BackgroundTask {
    param(
        [string]$Name,
        [array]$TaskArgs
    )
    if ($null -ne $global:CurrentJob -and (Get-Job -Id $global:CurrentJob.Id -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show("A task is already running. Cancel it first.", "Task Running", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return $null
    }
    
    $log = New-TaskLog $Name
    $global:CurrentTaskLog = $log
    Add-StatusUI $form $status "-> Starting task: $Name"
    
    $filterClass = $global:FilterSettings.Class
    $filterMfr = $global:FilterSettings.Manufacturer
    
    $job = Start-Job -Name $Name -ScriptBlock {
        param($taskName, $logPath, $innerArgs, $filterClass, $filterMfr)
        
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        function L($m) {
            $t = (Get-Date).ToString("s")
            Add-Content -Path $logPath -Value ("$t - $m")
        }
        
        try {
            switch ($taskName) {
                "DownloadAndInstallDrivers" {
                    # Get computer model
                    $computerModel = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
                    L "Detected computer model: $computerModel"
                    
                    # Load Nextcloud URLs
                    $modelsFileUrl = "https://nc.rikor.com/index.php/s/BfBKYyW9HdoFfz9/download"
                    $nextcloudUrls = @{}
                    
                    try {
                        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                        $webClient = New-Object System.Net.WebClient
                        $webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                        
                        $modelsJsonString = $webClient.DownloadString($modelsFileUrl)
                        $tempObj = $modelsJsonString | ConvertFrom-Json
                        $nextcloudUrls = @{}
                        
                        foreach($key in $tempObj.PSObject.Properties.Name) { 
                            $val = $tempObj.$key
                            if ($val -is [string]) {
                                $nextcloudUrls[$key] = $val.Trim().Trim('`"').Trim()
                            } elseif ($val.url) {
                                $nextcloudUrls[$key] = $val.url.Trim().Trim('`"').Trim()
                            }
                        }
                        L "Loaded models mapping from online file"
                    } catch {
                        L "Error loading online models.json: $_"
                    }
                    
                    $zipUrl = $null
                    $rikorServerAvailable = $false
                    
                    if ($nextcloudUrls.ContainsKey($computerModel)) {
                        $zipUrl = $nextcloudUrls[$computerModel]
                        L "Found driver URL for $computerModel"
                        
                        try {
                            $request = [System.Net.WebRequest]::Create($zipUrl)
                            $request.Method = "HEAD"
                            $request.Timeout = 15000
                            $request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                            $response = $request.GetResponse()
                            if ($response.StatusCode -eq 200) { $rikorServerAvailable = $true }
                            $response.Close()
                        } catch {
                            L "[INFO] Rikor server unavailable: $_"
                        }
                    } else {
                        L "Model '$computerModel' not in predefined list."
                    }
                    
                    function Install-FromMicrosoftUpdate {
                        param($LogFunction)
                        
                        & $LogFunction "[INFO] Starting Microsoft Update check..."
                        try {
                            $UpdateSession = New-Object -ComObject Microsoft.Update.Session
                            $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
                            $UpdateSearcher.ServerSelection = 0 # ssDefault
                            
                            & $LogFunction "Searching for driver updates..."
                            $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver'")
                            
                            if ($SearchResult.Updates.Count -eq 0) {
                                & $LogFunction "No driver updates found via Microsoft Update."
                            } else {
                                & $LogFunction "Found $($SearchResult.Updates.Count) driver update(s)."
                                
                                $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
                                foreach ($Update in $SearchResult.Updates) {
                                    if ($Update.EulaAccepted -eq $false) { $Update.AcceptEula() }
                                    $UpdatesToDownload.Add($Update) | Out-Null
                                }
                                
                                & $LogFunction "Downloading updates..."
                                $Downloader = $UpdateSession.CreateUpdateDownloader()
                                $Downloader.Updates = $UpdatesToDownload
                                $Downloader.Download()
                                
                                & $LogFunction "Installing updates..."
                                $Installer = $UpdateSession.CreateUpdateInstaller()
                                
                                $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                                foreach ($Update in $SearchResult.Updates) {
                                    if ($Update.IsDownloaded) { $UpdatesToInstall.Add($Update) | Out-Null }
                                }
                                
                                if ($UpdatesToInstall.Count -gt 0) {
                                    $Installer.Updates = $UpdatesToInstall
                                    $InstallResult = $Installer.Install()
                                    & $LogFunction "Installation finished. Result: $($InstallResult.ResultCode). Reboot: $($InstallResult.RebootRequired)"
                                }
                            }
                        } catch {
                            & $LogFunction "[ERROR] Microsoft Update failed: $_"
                        }
                    }
                    
                    if (-not $rikorServerAvailable) {
                        L "[INFO] Rikor server unavailable or model unknown. Falling back to Microsoft Update..."
                        Install-FromMicrosoftUpdate -LogFunction $function:L
                        L "Completed"
                        return
                    }
                    
                    # Download from Rikor
                    $rikorInstallSuccess = $false
                    $tempDir = Join-Path $env:TEMP "RikorDriversTemp_$(Get-Date -Format 'yyyyMMddHHmmss')"
                    try {
                        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                        $zipPath = Join-Path $tempDir "drivers.zip"
                        
                        L "Downloading drivers from Rikor..."
                        
                        # Define helper function for download with progress
                        function Download-WithProgress {
                            param([string]$Url, [string]$Path, [string]$LogFile)
                            $wc = New-Object System.Net.WebClient
                            $wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                            
                            # State object to track progress and avoid excessive file writes (IO Locking)
                            $state = @{ LastPercent = -1 }
                            
                            # Use GetNewClosure to ensure $LogFile and $state are captured
                            $evt = {
                                param($sender, $e)
                                $total = $e.TotalBytesToReceive
                                
                                if ($total -gt 0) {
                                    $p = [math]::Round(($e.BytesReceived / $total) * 100, 0)
                                    
                                    # Only write to log if percentage CHANGED to prevent file locking from rapid writes
                                    if ($p -ne $state.LastPercent) {
                                        $state.LastPercent = $p
                                        
                                        $mb = [math]::Round($e.BytesReceived / 1MB, 1)
                                        $totalMb = [math]::Round($total / 1MB, 1)
                                        $t = (Get-Date).ToString("s")
                                        
                                        try {
                                            $msg = "{0} - DL_PROGRESS:{1}:{2} MB/{3} MB`r`n" -f $t, $p, $mb, $totalMb
                                            [System.IO.File]::AppendAllText($LogFile, $msg)
                                        } catch {}
                                    }
                                }
                            }.GetNewClosure()
                            
                            $wc.add_DownloadProgressChanged($evt)
                            $wc.DownloadFile($Url, $Path)
                            $wc.Dispose()
                        }
                        
                        L "Starting download from Rikor Server..."
                        
                        Download-WithProgress -Url $zipUrl -Path $zipPath -LogFile $logPath
                        L "Download completed."
                        
                        if (-not (Test-Path $zipPath) -or (Get-Item $zipPath).Length -eq 0) {
                            throw "Download failed (empty file)."
                        }
                        
                        $extractDir = Join-Path $tempDir "ExtractedDrivers"
                        L "Extracting archive..."
                        Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
                        L "Extraction completed."
                        
                        $infFiles = Get-ChildItem -Path $extractDir -Recurse -Include "*.inf" -ErrorAction SilentlyContinue
                        L "Found $($infFiles.Count) driver files."
                        
                        $count = 0
                        $total = $infFiles.Count
                        foreach ($inf in $infFiles) {
                            $count++
                            L "[$count/$total] Installing $($inf.Name)..."
                            $out = & pnputil.exe /add-driver $inf.FullName /install /force 2>&1
                            # Basic check
                            if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne 259) {
                                L "   Error: $out"
                            }
                        }
                        L "Installation process finished."
                        $rikorInstallSuccess = $true
                        
                    } catch {
                        L "[ERROR] Rikor installation failed: $_"
                        $rikorInstallSuccess = $false
                    } finally {
                        if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
                    }
                    
                    if (-not $rikorInstallSuccess) {
                        L "[INFO] Rikor installation was unsuccessful. Falling back to Microsoft Update..."
                        Install-FromMicrosoftUpdate -LogFunction $function:L
                    }
                    
                    L "Completed"
                }
                
                "BackupDrivers" {
                    $dest = $innerArgs[0]
                    L "Backing up drivers to: $dest"
                    if (!(Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }
                    & dism.exe /online /export-driver /destination:$dest 2>&1 | Out-Null
                    L "Backup completed."
                    L "Completed"
                }
                
                "InstallDrivers" {
                    $folder = $innerArgs[0]
                    L "Installing drivers from: $folder"
                    $infFiles = Get-ChildItem -Path $folder -Recurse -Include *.inf -ErrorAction SilentlyContinue
                    L "Found $($infFiles.Count) drivers."
                    $count = 0
                    foreach ($inf in $infFiles) {
                        $count++
                        L "[$count/$($infFiles.Count)] Installing $($inf.Name)..."
                        & pnputil.exe /add-driver $inf.FullName /install /force 2>&1 | Out-Null
                    }
                    L "Completed"
                }
            }
        } catch {
            L "Job Error: $_"
        }
    } -ArgumentList $Name, $log, $TaskArgs, $filterClass, $filterMfr
    
    $global:CurrentJob = $job
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
                $invoke = [System.Windows.Forms.MethodInvoker]{
                    $status.Clear()
                    $status.AppendText($text + "`r`n")
                    $status.ScrollToCaret()
                }
                $form.Invoke($invoke)
            }
            
            # Progress parsing
            # (Download progress bar removed by user request)
            
            # Overall Progress
            $contentStr = $lines -join "`n"
            $p = 0
            if ($contentStr -match "Downloading") { $p = 10 }
            if ($contentStr -match "Download completed") { 
                $p = 30
            }
            if ($contentStr -match "Extracting") { $p = 40 }
            if ($contentStr -match "Extraction completed") { $p = 50 }
            if ($contentStr -match "Installing") { $p = 60 }
            
            # Find last installation progress
            $instLine = $lines | Where-Object { $_ -match "\[(\d+)\/(\d+)\]" } | Select-Object -Last 1
            if ($instLine -and $instLine -match "\[(\d+)\/(\d+)\]") {
                $curr = [int]$matches[1]
                $tot = [int]$matches[2]
                if ($tot -gt 0) { $p = 60 + [int](($curr / $tot) * 40) }
            }

            if ($contentStr -match "Completed") { 
                $p = 100
            }
            if ($p -gt 0) { 
                $progress.Value = $p 
                $overallLabel.Text = "$p%"
            }
        }
        
        if ($null -ne $global:CurrentJob) {
            $jobState = (Get-Job -Id $global:CurrentJob.Id -ErrorAction SilentlyContinue).State
            if ($jobState -in @("Completed","Failed","Stopped")) {
                Start-Sleep -Milliseconds 200
                $finishedText = "=== Task finished: $jobState ==="
                $invoke = [System.Windows.Forms.MethodInvoker]{
                    $status.AppendText("`r`n$finishedText`r`n")
                    $status.ScrollToCaret()
                    $progress.Value = 100
                    $statusLabel.Text = "  Task completed: $jobState"
                }
                $form.Invoke($invoke)
                
                # History
                Add-HistoryEntry -TaskName $global:CurrentJob.Name -Status $jobState -Details "Task completed"
                
                Remove-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
                $global:CurrentJob = $null
                $timer.Stop()
            }
        }
    } catch {}
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
    
    $downloadProgressPanel.BackColor = $colors.Background
    $downloadProgressBorder.BackColor = $colors.Border
    $overallProgressPanel.BackColor = $colors.Background
    $overallProgressBorder.BackColor = $colors.Border
    
    $overallLabel.ForeColor = $colors.Text
    
    $statusBar.BackColor = $colors.StatusBar
    $statusLabel.BackColor = $colors.StatusBar
    $statusLabel.ForeColor = [System.Drawing.Color]::White
    $versionLabel.BackColor = $colors.StatusBar
    $versionLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)
    
    if ($Dark) { $menuToggleTheme.Text = "Light Mode" } else { $menuToggleTheme.Text = "Dark Mode" }
    Export-Settings
}

# -------------------------
# Dialog Functions
# -------------------------
function Show-HistoryDialog {
    $colors = Get-ThemeColors
    $historyForm = New-Object Windows.Forms.Form
    $historyForm.Text = "Update History"
    $historyForm.Size = '750,500'
    $historyForm.StartPosition = "CenterParent"
    $historyForm.BackColor = $colors.Background
    $historyForm.TopMost = $false
    
    $header = New-Object Windows.Forms.Label
    $header.Text = "Update History"
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
    $schedForm.Text = "Schedule Updates"
    $schedForm.Size = '420,300'
    $schedForm.StartPosition = "CenterParent"
    $schedForm.BackColor = $colors.Background
    $schedForm.TopMost = $false
    
    $header = New-Object Windows.Forms.Label
    $header.Text = "Schedule Updates"
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
    $cmbFreq.Location = '110,8'
    $cmbFreq.Size = '180,28'
    $cmbFreq.DropDownStyle = 'DropDownList'
    $cmbFreq.Items.AddRange(@("Daily", "Weekly", "Monthly"))
    $cmbFreq.SelectedIndex = 0
    $contentPanel.Controls.Add($cmbFreq)
    
    $lblTime = New-Object Windows.Forms.Label
    $lblTime.Text = "Time:"
    $lblTime.Location = '0,50'
    $lblTime.Size = '100,25'
    $lblTime.ForeColor = $colors.Text
    $contentPanel.Controls.Add($lblTime)
    
    $txtTime = New-Object Windows.Forms.TextBox
    $txtTime.Location = '110,48'
    $txtTime.Size = '180,28'
    $txtTime.Text = "03:00"
    $contentPanel.Controls.Add($txtTime)
    
    $existingTask = Get-ScheduledUpdateTask
    $lblStatus = New-Object Windows.Forms.Label
    $lblStatus.Location = '0,95'
    $lblStatus.Size = '350,25'
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
    $btnEnable.Text = "Enable"
    $btnEnable.Location = '0,0'
    $btnEnable.Size = '110,36'
    $btnEnable.FlatStyle = 'Flat'
    $btnEnable.BackColor = $colors.Primary
    $btnEnable.ForeColor = [System.Drawing.Color]::White
    $btnEnable.Add_Click({
        if (Set-ScheduledUpdate -Frequency $cmbFreq.SelectedItem -Time $txtTime.Text) {
            $lblStatus.Text = "Scheduled task created successfully!"
            $lblStatus.ForeColor = $colors.Success
        }
    })
    $btnPanel.Controls.Add($btnEnable)
    
    $btnRemove = New-Object Windows.Forms.Button
    $btnRemove.Text = "Remove"
    $btnRemove.Location = '120,0'
    $btnRemove.Size = '120,36'
    $btnRemove.FlatStyle = 'Flat'
    $btnRemove.BackColor = $colors.Secondary
    $btnRemove.ForeColor = $colors.Text
    $btnRemove.Add_Click({
        Remove-ScheduledUpdate
        $lblStatus.Text = "Scheduled task removed."
        $lblStatus.ForeColor = $colors.TextSecondary
    })
    $btnPanel.Controls.Add($btnRemove)
    
    $btnClose = New-Object Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = '250,0'
    $btnClose.Size = '100,36'
    $btnClose.FlatStyle = 'Flat'
    $btnClose.BackColor = $colors.Secondary
    $btnClose.ForeColor = $colors.Text
    $btnClose.Add_Click({ $schedForm.Close() })
    $btnPanel.Controls.Add($btnClose)
    
    $schedForm.ShowDialog($form) | Out-Null
}

function Show-FiltersDialog {
    $colors = Get-ThemeColors
    $filterForm = New-Object Windows.Forms.Form
    $filterForm.Text = "Driver Filters"
    $filterForm.Size = '420,260'
    $filterForm.StartPosition = "CenterParent"
    $filterForm.BackColor = $colors.Background
    $filterForm.TopMost = $false
    
    $header = New-Object Windows.Forms.Label
    $header.Text = "Driver Filters"
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
    $lblClass.Text = "Filter by Class:"
    $lblClass.Location = '0,10'
    $lblClass.Size = '160,25'
    $lblClass.ForeColor = $colors.Text
    $contentPanel.Controls.Add($lblClass)
    
    $txtClass = New-Object Windows.Forms.TextBox
    $txtClass.Location = '170,8'
    $txtClass.Size = '180,28'
    $txtClass.Text = $global:FilterSettings.Class
    $contentPanel.Controls.Add($txtClass)
    
    $lblMfr = New-Object Windows.Forms.Label
    $lblMfr.Text = "Filter by Manufacturer:"
    $lblMfr.Location = '0,50'
    $lblMfr.Size = '160,25'
    $lblMfr.ForeColor = $colors.Text
    $contentPanel.Controls.Add($lblMfr)
    
    $txtMfr = New-Object Windows.Forms.TextBox
    $txtMfr.Location = '170,48'
    $txtMfr.Size = '180,28'
    $txtMfr.Text = $global:FilterSettings.Manufacturer
    $contentPanel.Controls.Add($txtMfr)
    
    $btnPanel = New-Object Windows.Forms.Panel
    $btnPanel.Location = '0,100'
    $btnPanel.Size = '360,40'
    $contentPanel.Controls.Add($btnPanel)
    
    $btnApply = New-Object Windows.Forms.Button
    $btnApply.Text = "Apply"
    $btnApply.Location = '0,0'
    $btnApply.Size = '110,36'
    $btnApply.FlatStyle = 'Flat'
    $btnApply.BackColor = $colors.Primary
    $btnApply.ForeColor = [System.Drawing.Color]::White
    $btnApply.Add_Click({
        $global:FilterSettings.Class = $txtClass.Text
        $global:FilterSettings.Manufacturer = $txtMfr.Text
        Export-Settings
        Add-StatusUI $form $status "Filter applied."
        $filterForm.Close()
    })
    $btnPanel.Controls.Add($btnApply)
    
    $btnClear = New-Object Windows.Forms.Button
    $btnClear.Text = "Clear"
    $btnClear.Location = '120,0'
    $btnClear.Size = '100,36'
    $btnClear.FlatStyle = 'Flat'
    $btnClear.BackColor = $colors.Secondary
    $btnClear.ForeColor = $colors.Text
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
    $btnClose.Text = "Close"
    $btnClose.Location = '230,0'
    $btnClose.Size = '100,36'
    $btnClose.FlatStyle = 'Flat'
    $btnClose.BackColor = $colors.Secondary
    $btnClose.ForeColor = $colors.Text
    $btnClose.Add_Click({ $filterForm.Close() })
    $btnPanel.Controls.Add($btnClose)
    
    $filterForm.ShowDialog($form) | Out-Null
}

function Show-SettingsDialog {
    $colors = Get-ThemeColors
    $settingsForm = New-Object Windows.Forms.Form
    $settingsForm.Text = "Settings"
    $settingsForm.Size = '500,340'
    $settingsForm.StartPosition = "CenterParent"
    $settingsForm.BackColor = $colors.Background
    $settingsForm.TopMost = $false
    
    $header = New-Object Windows.Forms.Label
    $header.Text = "Settings"
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
    
    $lblProxySection = New-Object Windows.Forms.Label
    $lblProxySection.Text = "Network Proxy"
    $lblProxySection.Location = '0,5'
    $lblProxySection.Size = '150,20'
    $lblProxySection.ForeColor = $colors.Primary
    $contentPanel.Controls.Add($lblProxySection)
    
    $lblProxy = New-Object Windows.Forms.Label
    $lblProxy.Text = "Proxy Address:"
    $lblProxy.Location = '0,35'
    $lblProxy.Size = '130,25'
    $lblProxy.ForeColor = $colors.Text
    $contentPanel.Controls.Add($lblProxy)
    
    $txtProxy = New-Object Windows.Forms.TextBox
    $txtProxy.Location = '140,33'
    $txtProxy.Size = '220,28'
    $txtProxy.Text = $global:ProxySettings.Address
    $contentPanel.Controls.Add($txtProxy)
    
    $chkProxy = New-Object Windows.Forms.CheckBox
    $chkProxy.Text = "Enable"
    $chkProxy.Location = '370,33'
    $chkProxy.Size = '80,25'
    $chkProxy.Checked = $global:ProxySettings.Enabled
    $chkProxy.ForeColor = $colors.Text
    $contentPanel.Controls.Add($chkProxy)
    
    $btnApplyProxy = New-Object Windows.Forms.Button
    $btnApplyProxy.Text = "Apply"
    $btnApplyProxy.Location = '140,70'
    $btnApplyProxy.Size = '110,34'
    $btnApplyProxy.FlatStyle = 'Flat'
    $btnApplyProxy.BackColor = $colors.Primary
    $btnApplyProxy.ForeColor = [System.Drawing.Color]::White
    $btnApplyProxy.Add_Click({
        Set-ProxySettings -ProxyAddr $txtProxy.Text -Enable $chkProxy.Checked
        Add-StatusUI $form $status "Proxy settings updated."
    })
    $contentPanel.Controls.Add($btnApplyProxy)
    
    $separator = New-Object Windows.Forms.Label
    $separator.Location = '0,120'
    $separator.Size = '440,1'
    $separator.BackColor = $colors.Border
    $contentPanel.Controls.Add($separator)
    
    $lblInfoSection = New-Object Windows.Forms.Label
    $lblInfoSection.Text = "Application Info"
    $lblInfoSection.Location = '0,130'
    $lblInfoSection.Size = '150,20'
    $lblInfoSection.ForeColor = $colors.Primary
    $contentPanel.Controls.Add($lblInfoSection)
    
    $lblInfo = New-Object Windows.Forms.Label
    $lblInfo.Text = "Logs: $LogBase"
    $lblInfo.Location = '0,155'
    $lblInfo.Size = '440,20'
    $lblInfo.ForeColor = $colors.TextSecondary
    $contentPanel.Controls.Add($lblInfo)
    
    $btnClose = New-Object Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = '170,210'
    $btnClose.Size = '110,36'
    $btnClose.FlatStyle = 'Flat'
    $btnClose.BackColor = $colors.Secondary
    $btnClose.ForeColor = $colors.Text
    $btnClose.Add_Click({ $settingsForm.Close() })
    $contentPanel.Controls.Add($btnClose)
    
    $settingsForm.ShowDialog($form) | Out-Null
}

# -------------------------
# Action Functions
# -------------------------
function Invoke-DownloadAndInstallDrivers {
    $status.Clear()
    $progress.Value = 0
    $statusLabel.Text = "  Downloading and installing drivers from Rikor..."
    Start-BackgroundTask -Name "DownloadAndInstallDrivers" -TaskArgs @()
}

function Invoke-BackupDrivers {
    $status.Clear()
    $progress.Value = 0
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select folder to save driver backup"
    $fbd.ShowNewFolderButton = $true
    if ($fbd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $dest = $fbd.SelectedPath
        $statusLabel.Text = "  Backing up drivers..."
        Start-BackgroundTask -Name "BackupDrivers" -TaskArgs @($dest)
    } else {
        Add-StatusUI $form $status "Backup canceled."
    }
}

function Invoke-InstallDrivers {
    $status.Clear()
    $progress.Value = 0
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select folder containing driver .inf files"
    $fbd.ShowNewFolderButton = $false
    if ($fbd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $folder = $fbd.SelectedPath
        $statusLabel.Text = "  Installing drivers..."
        Start-BackgroundTask -Name "InstallDrivers" -TaskArgs @($folder)
    } else {
        Add-StatusUI $form $status "Install canceled."
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
            Add-StatusUI $form $status "Task cancelled by user."
        } else {
            Add-StatusUI $form $status "No running task to cancel."
        }
    } catch {
        Add-StatusUI $form $status "Cancel error: $_"
    }
}

function Invoke-OpenLogs {
    if (Test-Path $LogBase) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"$LogBase`""
    } else {
        Add-StatusUI $form $status "Log folder missing."
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
        Add-StatusUI $form $status "System restore point created successfully!"
        $statusLabel.Text = "  Restore point created"
    } else {
        Add-StatusUI $form $status "Failed to create restore point."
        $statusLabel.Text = "  Restore point failed"
    }
    $progress.Value = 100
}

# -------------------------
# Event Handlers
# -------------------------
$btnDownloadAndInstall.Add_Click({ Invoke-DownloadAndInstallDrivers })
$btnBackup.Add_Click({ Invoke-BackupDrivers })
$btnInstall.Add_Click({ Invoke-InstallDrivers })
$btnCancel.Add_Click({ Invoke-CancelTask })

$menuDownloadAndInstall.Add_Click({ Invoke-DownloadAndInstallDrivers })
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

$form.Add_Resize({ Update-ButtonContainerPadding })
$form.Add_FormClosing({
    try {
        if ($null -ne $global:CurrentJob) {
            Stop-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
            Remove-Job -Id $global:CurrentJob.Id -Force -ErrorAction SilentlyContinue
        }
    } catch {}
})

Set-Theme -Dark $global:DarkModeEnabled
    
$form.Topmost = $false
$form.Add_Shown({
    $form.Activate()
    Update-ButtonContainerPadding
})

[void]$form.ShowDialog()
