# Driver Updater Rikor v 1.0
# Run as Administrator
# Usage: irm https://raw.githubusercontent.com/h4rd1ns0mn14/RikorDriver/refs/heads/main/RikorDriver.ps1 | iex
# Silent mode: .\updateAllDrivers.ps1 -Silent -Task "CheckDriverUpdates"
param(
[switch]$Silent,
[string]$Task = "",
[string]$FilterClass = "",
[string]$FilterManufacturer = ""
)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# -------------------------
# English-Only Language Support
# -------------------------
$global:Languages = @{
"en" = @{
AppTitle = "Rikor Driver Installer"
# BtnWU removed
BtnCheckUpdates = "Search and Install Rikor Drivers" # Updated text
BtnDownloadAndInstall = "Download and Install Rikor Drivers" # NEW: Localization for new button
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
}

# Get localized string
function Get-LocalizedString([string]$key) {
    return $global:Languages["en"][$key]
}
$global:CurrentLanguage = "en"
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
$global:FilterSettings = @{ Class = ""; Manufacturer = "" }
$global:DarkModeEnabled = $false
# -------------------------
# Settings Management
# -------------------------
function Import-Settings {
if (Test-Path $SettingsFile) {
try {
$settings = Get-Content -Path $SettingsFile -Raw | ConvertFrom-Json
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
Filters = $global:FilterSettings
DarkMode = $global:DarkModeEnabled
}
$settings | ConvertTo-Json -Depth 3 | Set-Content -Path $SettingsFile -Encoding UTF8
}
Import-Settings
# Apply command-line parameters
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
$entry = [PSCustomObject] @{
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
# CHANGED: Update scheduled task action to point to current script and new task name
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
$logFile = Join-Path $LogBase "$($Task)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Write-SilentLog($msg) {
$timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Add-Content -Path $logFile -Value "$timestamp - $msg"
}
Write-SilentLog "Starting silent mode: $Task"
switch ($Task) {
# REMOVED "WindowsUpdate" case
"CheckDriverUpdates" {
Write-SilentLog "Checking for driver updates and attempting to install..."
try {
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver'")
if ($SearchResult.Updates.Count -eq 0) {
Write-SilentLog "No driver updates available"
Add-HistoryEntry -TaskName "CheckDriverUpdates" -Status "Completed" -Details "No updates found"
} else {
Write-SilentLog "Found $($SearchResult.Updates.Count) driver update(s)"
foreach ($Update in $SearchResult.Updates) {
Write-SilentLog "  - $($Update.Title)"
}
Add-HistoryEntry -TaskName "CheckDriverUpdates" -Status "Completed" -Details "Found $($SearchResult.Updates.Count) updates"
}
} catch {
Write-SilentLog "Error: $_ "
Add-HistoryEntry -TaskName "CheckDriverUpdates" -Status "Failed" -Details $_.Exception.Message
}
}
# NEW: Add new task case for downloading and installing from ZIP
"DownloadAndInstallDrivers" { # NEW: Silent mode case for download and install
Write-SilentLog "Silent mode: Downloading and installing drivers from Rikor archive..."

# Get computer model to determine which driver pack to download
$computerModel = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
Write-SilentLog "Detected computer model: $computerModel"

# Load Nextcloud URLs from online JSON file
$modelsFileUrl = "https://nc.rikor.com/index.php/s/BfBKYyW9HdoFfz9/download"  # URL to the online models.json file
$nextcloudUrls = @{}

try {
# Attempt to download the models file
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

$modelsJsonString = $webClient.DownloadString($modelsFileUrl)
$tempObj = $modelsJsonString | ConvertFrom-Json
foreach($key in $tempObj.PSObject.Properties.Name) { $nextcloudUrls[$key] = $tempObj.$key }
} catch {
Write-SilentLog "Error downloading or parsing models from online file: $_ "
Write-SilentLog "Attempting to load from local models.json file as fallback"

# Fallback to local file if online retrieval fails
$modelsFilePath = Join-Path $PSScriptRoot "models.json"
if (Test-Path $modelsFilePath) {
try {
$tempObj = Get-Content -Path $modelsFilePath -Raw | ConvertFrom-Json
foreach($key in $tempObj.PSObject.Properties.Name) { $nextcloudUrls[$key] = $tempObj.$key }
} catch {
Write-SilentLog "Error loading models from local fallback file: $_ "
}
}
}

# Determine the appropriate URL based on model
$zipUrl = $null
if ($nextcloudUrls.ContainsKey($computerModel)) {
$zipUrl = $nextcloudUrls[$computerModel]
Write-SilentLog "Using Rikor Server for download"

# Check if Rikor server is available
$rikorServerAvailable = $false
try {
# Test connection to the URL using WebRequest with HEAD method (correct way to check availability)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$request = [System.Net.WebRequest]::Create($zipUrl)
$request.Method = "HEAD"
$request.Timeout = 15000  # 15 seconds timeout
$request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

$response = $request.GetResponse()
if ($response.StatusCode -eq 200) {
$rikorServerAvailable = $true
}
$response.Close()
} catch {
Write-SilentLog "[INFO] Rikor server is not accessible: $_ "
}
} else {
Write-SilentLog "Model '$computerModel' not in predefined list, checking Microsoft Update"
# Skip Nextcloud download and go directly to Microsoft Update
$rikorServerAvailable = $false
}

if (-not $rikorServerAvailable) {
Write-SilentLog "[INFO] Rikor server is not available. Checking for driver updates from Microsoft Update..."

try {
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
Write-SilentLog "Searching for available driver updates from Microsoft Update (this may take a few minutes)..."
$SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver'")

if ($SearchResult.Updates.Count -eq 0) {
Write-SilentLog "No driver updates available from Microsoft Update"
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = "DownloadAndInstallDrivers"
Status = "Completed"
Details = "No updates found (fallback to MS Update)"
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
} else {
Write-SilentLog "Found $($SearchResult.Updates.Count) driver update(s) available from Microsoft Update:"
Write-SilentLog ""
foreach ($Update in $SearchResult.Updates) {
Write-SilentLog "  - $($Update.Title)"
Write-SilentLog "    Size: $([math]::Round($Update.MaxDownloadSize / 1MB, 2)) MB"
}
Write-SilentLog ""
Write-SilentLog "Downloading and installing driver updates from Microsoft Update..."

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
Write-SilentLog "Downloading $($UpdatesToDownload.Count) update(s) from Microsoft Update..."
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
Write-SilentLog "Installing $($UpdatesToInstall.Count) update(s) from Microsoft Update..."
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
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = "DownloadAndInstallDrivers"
Status = "Completed"
Details = "Installed $successCount/$($UpdatesToInstall.Count) updates via MS Update"
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
} else {
Write-SilentLog "[WARNING] No updates were downloaded successfully from Microsoft Update"
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = "DownloadAndInstallDrivers"
Status = "Completed"
Details = "Download failed for all updates from MS Update"
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
}
}
} catch {
Write-SilentLog "[ERROR] Failed to download/install driver updates from Microsoft Update: $_ "
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = "DownloadAndInstallDrivers"
Status = "Failed"
Details = $_.Exception.Message
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
}
Write-SilentLog "Completed"
return
}

# Validate the URL
if (-not $zipUrl -or $zipUrl -eq "https://drive.google.com/uc?export=download&id=14_iaT8zdS800GpL76CSVb5vBQN7whZ8w") {
Write-SilentLog "[ERROR] Public ZIP download URL is not configured. Please edit the script."
Write-SilentLog "Replace 'https://drive.google.com/uc?export=download&id=14_iaT8zdS800GpL76CSVb5vBQN7whZ8w' with the actual link."
Write-SilentLog "Completed"
return
}

# Create temp directory
$tempDir = Join-Path $env:TEMP "RikorDriversTemp_$(Get-Date -Format 'yyyyMMddHHmmss')"
try {
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
Write-SilentLog "Created temporary directory: $tempDir"

# Define zip file path
$zipPath = Join-Path $tempDir "drivers.zip"

# Download ZIP
Write-SilentLog "Starting download of drivers archive..."
try {
# Force TLS 1.2 for Nextcloud compatibility
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Use custom download function with progress for reliable download from Nextcloud
Write-SilentLog "Starting download with progress tracking..."
try {
# Define helper function for download with progress
function Download-WithProgress {
param(
[string]$Url,
[string]$Path,
[string]$Activity = "Downloading"
)

# Create WebClient with events for progress tracking
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

# Create temporary file for download
$tempPath = $Path + ".tmp"

# Register progress event handler
$webClient.add_DownloadProgressChanged({
param($sender, $event)

# If server returned file size (TotalBytesToReceive > 0)
if ($event.TotalBytesToReceive -gt 0) {
$percent = [math]::Round(($event.BytesReceived / $event.TotalBytesToReceive) * 100, 2)
$receivedMB = [math]::Round($event.BytesReceived / 1MB, 2)
$totalMB = [math]::Round($event.TotalBytesToReceive / 1MB, 2)

Write-Progress -Activity $Activity -Status "$percent% Complete ($receivedMB/$totalMB MB)" -PercentComplete $percent
}
else {
# If file size is unknown (typical for Nextcloud dynamic links), show downloaded amount only
$receivedMB = [math]::Round($event.BytesReceived / 1MB, 2)
Write-Progress -Activity $Activity -Status "Downloaded: $receivedMB MB (Size unknown...)" -PercentComplete 0
}
})

try {
$webClient.DownloadFile($Url, $tempPath)

# Move temp file to final location when complete
Move-Item -Path $tempPath -Destination $Path -Force
Write-SilentLog "Download completed: $Path"
} finally {
# Clean up temp file if it still exists
if (Test-Path $tempPath) {
Remove-Item $tempPath -ErrorAction SilentlyContinue
}
# Clear progress bar
Write-Progress -Activity $Activity -Completed
}
}

Download-WithProgress -Url $zipUrl -Path $zipPath -Activity "Downloading Drivers Archive"
Write-SilentLog "Download completed successfully: $zipPath"
} catch {
Write-SilentLog "[ERROR] Download failed with progress tracking: $_ "
# Fallback to basic download
try {
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
$webClient.DownloadFile($zipUrl, $zipPath)
Write-SilentLog "Fallback download completed: $zipPath"
} catch {
Write-SilentLog "[ERROR] Fallback download also failed: $_ "
throw
}
}
Write-SilentLog "Download completed to: $zipPath"
} catch {
Write-SilentLog "[ERROR] Failed to download ZIP: $_ "
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-SilentLog "Completed"
return
}

# Check if ZIP exists and is not empty
if (-not (Test-Path $zipPath) -or (Get-Item $zipPath).Length -eq 0) {
Write-SilentLog "[ERROR] Downloaded ZIP file is missing or empty."
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-SilentLog "Completed"
return
}

# Extract ZIP
$extractDir = Join-Path $tempDir "ExtractedDrivers"
Write-SilentLog "Extracting ZIP archive to: $extractDir"
try {
Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
Write-SilentLog "Extraction completed."
} catch {
Write-SilentLog "[ERROR] Failed to extract ZIP: $_ "
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-SilentLog "Completed"
return
}

# Verify extracted content
if (-not (Test-Path $extractDir)) {
Write-SilentLog "[ERROR] Extraction directory '$extractDir' does not exist after extraction."
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-SilentLog "Completed"
return
}

$infFilesInZip = Get-ChildItem -Path $extractDir -Recurse -Include "*.inf" -ErrorAction SilentlyContinue
if ($infFilesInZip.Count -eq 0) {
Write-SilentLog "[WARNING] No .inf files found in the downloaded archive after extraction."
# Still proceed to InstallDrivers task, which will report its own error if needed
} else {
Write-SilentLog "Found $($infFilesInZip.Count) .inf file(s) in the archive."
}

# Prepare arguments for the InstallDrivers task logic
$installArgs = @($extractDir)

# --- Simulate calling the InstallDrivers logic directly within the same job ---
# This avoids creating another nested job and keeps logging consistent.
$folder = $installArgs[0]
Write-SilentLog "Installing drivers from extracted archive: $folder"
try {
if (-not (Test-Path $folder)) {
Write-SilentLog "[ERROR] Folder not found: $folder (This should not happen after extraction)"
Write-SilentLog "Completed"
return
}

# Find INF files recursively in the extracted folder
$infFiles = Get-ChildItem -Path $folder -Recurse -Include *.inf -ErrorAction SilentlyContinue

if ($infFiles.Count -eq 0) {
Write-SilentLog "[ERROR] No .inf driver files found in the extracted folder: $folder"
Write-SilentLog "Completed"
return
}

Write-SilentLog "Found $($infFiles.Count) driver file(s) in extracted folder"
Write-SilentLog "Installing drivers..."
Write-SilentLog ""
$successCount = 0
$failCount = 0
$current = 0
foreach ($inf in $infFiles) {
$current++
Write-SilentLog "[$current/$($infFiles.Count)] $($inf.Name)"
try {
# Use pnputil to add and install the driver
$out = & pnputil.exe /add-driver $inf.FullName /install /force 2>&1
$hasErrorInOutput = $false
$out | ForEach-Object {
if ($_ -match "(error|failed|fail|cannot find suitable|not compatible|not applicable)") {
Write-SilentLog "     $_ " # Log the error line
$hasErrorInOutput = $true
}
}

if ($hasErrorInOutput) {
$failCount++
Write-SilentLog "     -> Installation failed (error found in pnputil output)."
} elseif ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 259) {
# 0 = success, 259 = no more items (already installed or not applicable)
$successCount++
if ($LASTEXITCODE -eq 259) {
Write-SilentLog "     -> Driver is already up to date or no applicable device found (pnputil exit code: 259)."
} else {
Write-SilentLog "     -> Added and installed successfully."
}
} else {
# Non-zero exit code that isn't 259
$failCount++
Write-SilentLog "     -> Installation failed (pnputil exit code: $LASTEXITCODE)."
}
} catch {
$failCount++
Write-SilentLog "     -> Exception during installation: $_ "
}
Start-Sleep -Milliseconds 300
}
Write-SilentLog ""
Write-SilentLog "Installation from archive complete: $successCount successful, $failCount failed"
if ($successCount -gt 0) {
Write-SilentLog "Note: Reboot may be required for some drivers to take effect."
}
} catch {
Write-SilentLog "[ERROR] Installation process failed: $_ "
}
# --- End of simulated InstallDrivers logic ---

} catch {
Write-SilentLog "[ERROR] An unexpected error occurred during download/extraction: $_ "
} finally {
# Always attempt to clean up the temporary directory
Write-SilentLog "Cleaning up temporary directory: $tempDir"
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
if (Test-Path $tempDir) {
Write-SilentLog "[WARNING] Could not remove temporary directory: $tempDir"
} else {
Write-SilentLog "Temporary directory cleaned up successfully."
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
Write-SilentLog "[ERROR] Failed to scan drivers: $_ "
}
}
"BackupDrivers" {
$dest = $innerArgs[0]
Write-SilentLog "Backing up drivers to: $dest"
try {
if (!(Test-Path $dest)) {
New-Item -ItemType Directory -Path $dest -Force | Out-Null
}
Write-SilentLog "Exporting drivers (this may take several minutes)...
"
& dism.exe /online /export-driver /destination:$dest 2>&1 | Out-Null
$exportedCount = (Get-ChildItem -Path $dest -Recurse -Directory -ErrorAction SilentlyContinue).Count
if ($exportedCount -gt 0) {
Write-SilentLog "Backup completed: $exportedCount driver package(s) exported"
} else {
Write-SilentLog "Backup completed. Check destination folder for exported drivers."
}
} catch {
Write-SilentLog "[ERROR] Backup failed: $_ "
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
foreach ($inf in $infFiles) {
$current++
Write-SilentLog "[$current/$($infFiles.Count)] $($inf.Name)"
try {
$out = & pnputil.exe /add-driver $inf.FullName /install /force 2>&1
$hasError = $false
$out | ForEach-Object {
if ($_ -match "(error|failed|fail|cannot find suitable)" -and -not $hasError) {
Write-SilentLog "     $_ " # Log the error line
$hasError = $true
$failCount++
}
}
if (-not $hasError) {
# If no specific error was logged, check the exit code
if ($LASTEXITCODE -eq 0) {
$successCount++
Write-SilentLog "     -> Added and installed successfully." # Optional verbose log
} else {
# pnputil reported an error via exit code, even if not captured in output
$failCount++
Write-SilentLog "     -> Installation failed (pnputil exit code: $LASTEXITCODE)."
}
}
} catch {
$failCount++
Write-SilentLog "     -> Exception during installation: $_ "
}
Start-Sleep -Milliseconds 300
}
Write-SilentLog ""
Write-SilentLog "Installation complete: $successCount successful, $failCount failed"
if ($successCount -gt 0) {
Write-SilentLog "Note: Reboot may be required for some drivers."
}
} catch {
Write-SilentLog "[ERROR] Installation failed: $_ "
}
Write-SilentLog "Completed"
}
default {
Write-SilentLog "Unknown task: $Task"
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = $Task
Status = "Failed"
Details = "Unknown task"
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
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
$menuFile.DropDownItems.AddRange( @($menuOpenLogs, $menuSeparator1, $menuExit))
# Actions Menu
$menuActions = New-Object Windows.Forms.ToolStripMenuItem
$menuActions.Text = "&Actions"
# REMOVED menuWU
# REMOVED menuCheckUpdates
# NEW: Add new menu item for download and install
$menuDownloadAndInstall = New-Object Windows.Forms.ToolStripMenuItem
$menuDownloadAndInstall.Text = (Get-LocalizedString "BtnDownloadAndInstall")
$menuDownloadAndInstall.ShortcutKeys = [System.Windows.Forms.Keys]::F6 # Assign F6 to the combined button
# Insert it near the top of the actions menu
$menuActions.DropDownItems.AddRange( @($menuDownloadAndInstall)) # Adjust order as needed
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
$menuActions.DropDownItems.AddRange( @($menuScan, $menuSeparator2, $menuBackup, $menuInstall, $menuSeparator3, $menuCancel))
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
$menuTools.DropDownItems.AddRange( @($menuRestorePoint, $menuSchedule, $menuFilters, $menuSeparator4, $menuHistory))
# View Menu
$menuView = New-Object Windows.Forms.ToolStripMenuItem
$menuView.Text = "&View"
$menuToggleTheme = New-Object Windows.Forms.ToolStripMenuItem
$menuToggleTheme.Text = (Get-LocalizedString "BtnDarkMode")
$menuToggleTheme.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::T
$menuView.DropDownItems.AddRange( @($menuToggleTheme))
# Settings Menu
$menuSettingsTop = New-Object Windows.Forms.ToolStripMenuItem
$menuSettingsTop.Text = "&Settings"
$menuSettingsTop.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Oemcomma
$menuStrip.Items.AddRange( @($menuFile, $menuActions, $menuTools, $menuView, $menuSettingsTop))
# -------------------------
# Toolbar Panel
# -------------------------
$toolbarPanel = New-Object Windows.Forms.Panel
$toolbarPanel.Dock = 'Top'
$toolbarPanel.Height = 56
# Note: toolbarPanel and menuStrip are added later for correct dock order
# FlowLayoutPanel for centered buttons
$buttonContainer = New-Object Windows.Forms.FlowLayoutPanel
$buttonContainer.Dock = 'Fill'
$buttonContainer.FlowDirection = 'LeftToRight'
$buttonContainer.WrapContents = $false
$buttonContainer.AutoSize = $false
$buttonContainer.Padding = '0,8,0,8'
$toolbarPanel.Controls.Add($buttonContainer)
# Function to center buttons when form resizes
function Update-ButtonContainerPadding {
# NEW: Include width of new combined button (only one button remains)
$totalButtonWidth = 160 + 155 + 120 + 140 + 110 + (4 * 12)  # buttons + gaps (Removed width for btnWU and old btnCheckUpdates)
$availableWidth = $toolbarPanel.ClientSize.Width
$leftPadding = [Math]::Max(0, [int](($availableWidth - $totalButtonWidth) / 2))
$buttonContainer.Padding = "$leftPadding,8,0,8"
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
# Create toolbar buttons with margins for spacing
# REMOVED $btnWU
# REMOVED $btnCheckUpdates
# NEW: Create button for download and install (only one button remains)
$btnDownloadAndInstall = New-ModernButton -Text (Get-LocalizedString "BtnDownloadAndInstall") -Width 160 -Primary $true
$btnDownloadAndInstall.Margin = '0,0,12,0'
$btnScan = New-ModernButton -Text (Get-LocalizedString "BtnScan") -Width 155 -Primary $true
$btnScan.Margin = '0,0,12,0'
$btnBackup = New-ModernButton -Text (Get-LocalizedString "BtnBackup") -Width 120
$btnBackup.Margin = '0,0,12,0'
$btnInstall = New-ModernButton -Text (Get-LocalizedString "BtnInstall") -Width 140
$btnInstall.Margin = '0,0,12,0'
$btnCancel = New-ModernButton -Text (Get-LocalizedString "BtnCancel") -Width 110
$btnCancel.Margin = '0,0,0,0'
# Add buttons to centered container (Only the combined button remains)
$buttonContainer.Controls.AddRange( @($btnDownloadAndInstall, $btnScan, $btnBackup, $btnInstall, $btnCancel))
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
# Force TLS 1.2 for Nextcloud compatibility
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Logging function
function L($m) {
    $t = (Get-Date).ToString("s")
    Add-Content -Path $logPath -Value ("$t - $m")
}
try {
# Execute task based on task name
switch ($taskName) {
# REMOVED "WindowsUpdate" case
# REMOVED "CheckDriverUpdates" case
"DownloadAndInstallDrivers" { # NEW: Combined task name
    # FORCE TLS 1.2 for Nextcloud compatibility
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    # Get computer model to determine which driver pack to download
    $computerModel = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
    L "Detected computer model: $computerModel"
    
    # Load Nextcloud URLs from online JSON file
    $modelsFileUrl = "https://nc.rikor.com/index.php/s/BfBKYyW9HdoFfz9/download"  # URL to the online models.json file
    $nextcloudUrls = @{}
    
try {
        # Attempt to download the models file
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        $modelsJsonString = $webClient.DownloadString($modelsFileUrl)
        $tempObj = $modelsJsonString | ConvertFrom-Json
        $nextcloudUrls = @{}
        foreach($key in $tempObj.PSObject.Properties.Name) { $nextcloudUrls[$key] = $tempObj.$key }
        L "Loaded models mapping from online file"
    } catch {
        L "Error downloading or parsing models from online file: $_ "
        L "Attempting to load from local models.json file as fallback"
        
        # Fallback to local file if online retrieval fails
        $modelsFilePath = Join-Path $PSScriptRoot "models.json"
        if (Test-Path $modelsFilePath) {
try {
$tempObj = Get-Content -Path $modelsFilePath -Raw | ConvertFrom-Json
$nextcloudUrls = @{}
foreach($key in $tempObj.PSObject.Properties.Name) { $nextcloudUrls[$key] = $tempObj.$key }
L "Loaded models mapping from local fallback file"
} catch {
L "Error loading models from local fallback file: $_ "
}
}
}

# Determine the appropriate URL based on model
$zipUrl = $null
if ($nextcloudUrls.ContainsKey($computerModel)) {
$zipUrl = $nextcloudUrls[$computerModel]
L "Using Rikor Server for download"

# Check if Rikor server is available
$rikorServerAvailable = $false
try {
# Test connection to the URL using WebRequest with HEAD method (correct way to check availability)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$request = [System.Net.WebRequest]::Create($zipUrl)
$request.Method = "HEAD"
$request.Timeout = 15000  # 15 seconds timeout
$request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

$response = $request.GetResponse()
if ($response.StatusCode -eq 200) {
$rikorServerAvailable = $true
}
$response.Close()
} catch {
L "[INFO] Rikor server is not accessible: $_ "
}
} else {
L "Model '$computerModel' not in predefined list, checking Microsoft Update"
# Skip Nextcloud download and go directly to Microsoft Update
$rikorServerAvailable = $false
}


if (-not $rikorServerAvailable) {
L "[INFO] Rikor server is not available. Checking for driver updates from Microsoft Update..."

try {
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
L "Searching for available driver updates from Microsoft Update (this may take a few minutes)..."
$SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver'")

if ($SearchResult.Updates.Count -eq 0) {
L "No driver updates available from Microsoft Update"
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = "DownloadAndInstallDrivers"
Status = "Completed"
Details = "No updates found (fallback to MS Update)"
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
} else {
L "Found $($SearchResult.Updates.Count) driver update(s) available from Microsoft Update:"
L ""
foreach ($Update in $SearchResult.Updates) {
L "  - $($Update.Title)"
L "    Size: $([math]::Round($Update.MaxDownloadSize / 1MB, 2)) MB"
}
L ""
L "Downloading and installing driver updates from Microsoft Update..."

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
L "Downloading $($UpdatesToDownload.Count) update(s) from Microsoft Update..."
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
L "Installing $($UpdatesToInstall.Count) update(s) from Microsoft Update..."
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
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = "DownloadAndInstallDrivers"
Status = "Completed"
Details = "Installed $successCount/$($UpdatesToInstall.Count) updates via MS Update"
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
} else {
L "[WARNING] No updates were downloaded successfully from Microsoft Update"
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = "DownloadAndInstallDrivers"
Status = "Completed"
Details = "Download failed for all updates from MS Update"
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
}
}
} catch {
L "[ERROR] Failed to download/install driver updates from Microsoft Update: $_ "
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = "DownloadAndInstallDrivers"
Status = "Failed"
Details = $_.Exception.Message
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
}
L "Completed"
return
}

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
L "Starting download of drivers archive..."
try {
# Force TLS 1.2 for Nextcloud compatibility
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Use custom download function with progress for reliable download from Nextcloud
L "Starting download with progress tracking..."
try {
# Define helper function for download with progress
function Download-WithProgress {
param(
[string]$Url,
[string]$Path,
[string]$Activity = "Downloading"
)

# Create WebClient with events for progress tracking
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

# Create temporary file for download
$tempPath = $Path + ".tmp"

# Register progress event handler
$webClient.add_DownloadProgressChanged({
param($sender, $event)

# If server returned file size (TotalBytesToReceive > 0)
if ($event.TotalBytesToReceive -gt 0) {
$percent = [math]::Round(($event.BytesReceived / $event.TotalBytesToReceive) * 100, 2)
$receivedMB = [math]::Round($event.BytesReceived / 1MB, 2)
$totalMB = [math]::Round($event.TotalBytesToReceive / 1MB, 2)

Write-Progress -Activity $Activity -Status "$percent% Complete ($receivedMB/$totalMB MB)" -PercentComplete $percent
}
else {
# If file size is unknown (typical for Nextcloud dynamic links), show downloaded amount only
$receivedMB = [math]::Round($event.BytesReceived / 1MB, 2)
Write-Progress -Activity $Activity -Status "Downloaded: $receivedMB MB (Size unknown...)" -PercentComplete 0
}
})

try {
$webClient.DownloadFile($Url, $tempPath)

# Move temp file to final location when complete
Move-Item -Path $tempPath -Destination $Path -Force
Write-SilentLog "Download completed: $Path"
} finally {
# Clean up temp file if it still exists
if (Test-Path $tempPath) {
Remove-Item $tempPath -ErrorAction SilentlyContinue
}
# Clear progress bar
Write-Progress -Activity $Activity -Completed
}
}

Download-WithProgress -Url $zipUrl -Path $zipPath -Activity "Downloading Drivers Archive"
L "Download completed successfully: $zipPath"
} catch {
L "[ERROR] Download failed with progress tracking: $_ "
# Fallback to basic download
try {
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$webClient = New-Object System.Net.WebClient
$webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
$webClient.DownloadFile($zipUrl, $zipPath)
L "Fallback download completed: $zipPath"
} catch {
L "[ERROR] Fallback download also failed: $_ "
throw
}
}
L "Download completed to: $zipPath"
} catch {
L "[ERROR] Failed to download ZIP: $_ "
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
L "[ERROR] Failed to extract ZIP: $_ "
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
L "Installing drivers from extracted archive: $folder"
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
$hasErrorInOutput = $false
$out | ForEach-Object {
if ($_ -match "(error|failed|fail|cannot find suitable|not compatible|not applicable)") {
L "     $_ " # Log the error line
$hasErrorInOutput = $true
}
}

if ($hasErrorInOutput) {
$failCount++
L "     -> Installation failed (error found in pnputil output)."
} elseif ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 259) {
# 0 = success, 259 = no more items (already installed or not applicable)
$successCount++
if ($LASTEXITCODE -eq 259) {
L "     -> Driver is already up to date or no applicable device found (pnputil exit code: 259)."
} else {
L "     -> Added and installed successfully."
}
} else {
# Non-zero exit code that isn't 259
$failCount++
L "     -> Installation failed (pnputil exit code: $LASTEXITCODE)."
}
} catch {
$failCount++
L "     -> Exception during installation: $_ "
}
Start-Sleep -Milliseconds 300
}
L ""
L "Installation from archive complete: $successCount successful, $failCount failed"
if ($successCount -gt 0) {
L "Note: Reboot may be required for some drivers to take effect."
}
} catch {
L "[ERROR] Installation process failed: $_ "
}
# --- End of simulated InstallDrivers logic ---

} catch {
L "[ERROR] An unexpected error occurred during download/extraction: $_ "
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
L "[ERROR] Failed to scan drivers: $_ "
}
}
"BackupDrivers" {
$dest = $innerArgs[0]
L "Backing up drivers to: $dest"
try {
if (!(Test-Path $dest)) {
New-Item -ItemType Directory -Path $dest -Force | Out-Null
}
L "Exporting drivers (this may take several minutes)...
"
& dism.exe /online /export-driver /destination:$dest 2>&1 | Out-Null
$exportedCount = (Get-ChildItem -Path $dest -Recurse -Directory -ErrorAction SilentlyContinue).Count
if ($exportedCount -gt 0) {
L "Backup completed: $exportedCount driver package(s) exported"
} else {
L "Backup completed. Check destination folder for exported drivers."
}
} catch {
L "[ERROR] Backup failed: $_ "
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
L "     $_ " # Log the error line
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
L "     -> Installation failed (pnputil exit code: $LASTEXITCODE)."
}
}
} catch {
$failCount++
L "     -> Exception during installation: $_ "
}
Start-Sleep -Milliseconds 300
}
L ""
L "Installation complete: $successCount successful, $failCount failed"
if ($successCount -gt 0) {
L "Note: Reboot may be required for some drivers."
}
} catch {
L "[ERROR] Installation failed: $_ "
}
L "Completed"
}
default {
L "ERROR: Unknown task name: $taskName"
}
}
} catch {
L "ERROR in job: $_ "
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
# NEW: Enhanced progress detection for download with percentage tracking
if ($contentStr -match "(\d+)% Complete") {
    $matches = [regex]::Matches($contentStr, "(\d+)% Complete")
    if ($matches.Count -gt 0) {
        $downloadPercent = [int]$matches[$matches.Count-1].Groups[1].Value
        # Map download progress (0-100%) to UI progress range (5-15% of total progress bar)
        $p = 5 + [int](($downloadPercent / 100) * 10)  # 10% of total progress for download phase
    }
}
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
# Direct history entry addition (instead of Add-HistoryEntry function which is not available in background job)
$historyFile = Join-Path $env:USERPROFILE "Documents\Rikor_DriverInstaller\UpdateHistory.json"
$history = @()
if (Test-Path $historyFile) {
try {
$content = Get-Content -Path $historyFile -Raw
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
$entry = [PSCustomObject] @{
Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
Task = $taskName
Status = $jobState
Details = "Task completed"
}
$history = @($entry) + $history
# Keep last 100 entries
if ($history.Count -gt 100) { $history = $history[0..99] }
$history | ConvertTo-Json -Depth 3 | Set-Content -Path $historyFile -Encoding UTF8
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
# REMOVED $btnWU.Text update
# REMOVED $btnCheckUpdates.Text update
# NEW: Update new button text (CORRECTED: Removed extra closing parenthesis)
$btnDownloadAndInstall.Text = Get-LocalizedString "BtnDownloadAndInstall"
$btnScan.Text = Get-LocalizedString "BtnScan"
$btnBackup.Text = Get-LocalizedString "BtnBackup"
$btnInstall.Text = Get-LocalizedString "BtnInstall"
$btnCancel.Text = Get-LocalizedString "BtnCancel"
# Menu items
$menuOpenLogs.Text = Get-LocalizedString "BtnOpenLogs"
# REMOVED $menuWU.Text update
# REMOVED $menuCheckUpdates.Text update
# NEW: Update new menu item text (CORRECTED: Removed extra closing parenthesis)
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
$cmbFreq.Items.AddRange( @((Get-LocalizedString "Daily"), (Get-LocalizedString "Weekly"), (Get-LocalizedString "Monthly")))
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
    # Info section$lblInfoSection = New-Object Windows.Forms.Label
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
# REMOVED Invoke-WindowsUpdate
# REMOVED Invoke-CheckDriverUpdates
# NEW: Define action function for new combined task
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
Add-StatusUI $form $status "Cancel error: $_ "
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
# REMOVED $btnWU.Add_Click
# REMOVED $btnCheckUpdates.Add_Click
# NEW: Add click handler for new combined button
$btnDownloadAndInstall.Add_Click({ Invoke-DownloadAndInstallDrivers })
$btnScan.Add_Click({ Invoke-ScanDrivers })
$btnBackup.Add_Click({ Invoke-BackupDrivers })
$btnInstall.Add_Click({ Invoke-InstallDrivers })
$btnCancel.Add_Click({ Invoke-CancelTask })
# -------------------------
# Menu handlers
# -------------------------
# REMOVED $menuWU.Add_Click
# REMOVED $menuCheckUpdates.Add_Click
# NEW: Add click handler for new combined menu item
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
