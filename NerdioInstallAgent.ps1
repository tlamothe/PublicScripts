# =========================
# CONFIG – EDIT THESE
# =========================
$BlobSasUrl   = 'https://<youraccount>.blob.core.windows.net/<container>/ResolveAgent.msi?<SAS>'
$ExpectedHash = '<optional SHA256 hash here, or leave empty>'
$InstallerIsMSI = $true   # set $false if yours is an .exe installer

# If your Resolve installer needs enrollment parameters, add them here.
# Example MSI properties (replace with your real values or leave blank):
$MsiProps = 'SITEID=xxxx ENROLLMENTKEY=yyyy'  # or '' if not needed

# For EXE installers, set silent args:
$ExeInstallArgs = '/quiet /norestart'         # adjust to vendor docs

# Optional vendor-specific uninstall fallback (only used if we can't find an uninstall string)
# For MSI: leave blank; for EXE provide a command line that runs silently
$ExeUninstallArgs = '/uninstall /quiet /norestart'

# =========================
# CONSTANTS
# =========================
$BaseDir    = 'C:\ProgramData\ResolveDeploy'
$LogsDir    = Join-Path $BaseDir 'Logs'
$Installer  = Join-Path $BaseDir 'ResolveAgent' + ($(if($InstallerIsMSI){'.msi'}else{'.exe'}))
$InstLog    = Join-Path $LogsDir 'install.log'
$UninstLog  = Join-Path $LogsDir 'uninstall.log'
$InstallPs1 = Join-Path $BaseDir 'Install-ResolveOnStartup.ps1'
$UninstPs1  = Join-Path $BaseDir 'Remove-ResolveOnShutdown.ps1'

New-Item -ItemType Directory -Force -Path $BaseDir, $LogsDir | Out-Null

# =========================
# WRITE: Install-ResolveOnStartup.ps1
# =========================
@"
param(
  [string]\$BlobUrl = '$BlobSasUrl',
  [string]\$ExpectedSha256 = '$ExpectedHash',
  [string]\$InstallerPath = '$Installer',
  [string]\$LogPath = '$InstLog',
  [bool]\$IsMSI = $InstallerIsMSI,
  [string]\$MsiProps = '$MsiProps',
  [string]\$ExeInstallArgs = '$ExeInstallArgs'
)

\$ErrorActionPreference = 'Stop'
Start-Transcript -Path \$LogPath -Append

function Write-Log(\$msg){ \$ts = (Get-Date).ToString('s'); Write-Output "[\$ts] \$msg" }

# Detect installed Resolve (broad but safe). Adjust DisplayName filter if needed.
function Get-ResolveUninstallInfo {
  \$keys = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )
  foreach(\$k in \$keys){
    Get-ItemProperty -Path \$k -ErrorAction SilentlyContinue |
      Where-Object { \$_.DisplayName -match 'Resolve|GoTo Resolve|LogMeIn.*Resolve' }
  }
}

# Idempotency: already installed?
\$installed = Get-ResolveUninstallInfo | Select-Object -First 1
if(\$installed){
  Write-Log "Resolve appears installed: \$((\$installed.DisplayName) -replace '\s+',' ')"
  Stop-Transcript
  exit 0
}

# Download (skip if already present & hash matches)
\$needsDownload = \$true
if(Test-Path \$InstallerPath -PathType Leaf -ErrorAction SilentlyContinue -and \$ExpectedSha256){
  try{
    \$h = (Get-FileHash -Path \$InstallerPath -Algorithm SHA256).Hash
    if(\$h -ieq \$ExpectedSha256){ \$needsDownload = \$false; Write-Log "Installer already cached & hash matches." }
    else { Write-Log "Cached installer hash mismatch; re-downloading." }
  }catch{}
}

if(\$needsDownload){
  Write-Log "Downloading installer from SAS URL..."
  if(Test-Path \$InstallerPath){ Remove-Item \$InstallerPath -Force -ErrorAction SilentlyContinue }
  try{
    # BITS is resilient on flaky networks
    Start-BitsTransfer -Source \$BlobUrl -Destination \$InstallerPath -Description 'ResolveInstaller' -ErrorAction Stop
  }catch{
    Write-Log "BITS failed: \$($_.Exception.Message). Falling back to Invoke-WebRequest."
    Invoke-WebRequest -UseBasicParsing -Uri \$BlobUrl -OutFile \$InstallerPath
  }
  if(\$ExpectedSha256){
    \$actual = (Get-FileHash -Path \$InstallerPath -Algorithm SHA256).Hash
    if(\$actual -ine \$ExpectedSha256){
      Write-Log "ERROR: SHA256 mismatch. Expected \$ExpectedSha256, got \$actual"
      throw "Installer hash mismatch"
    }
  }
}

# Install silently
Write-Log "Installing Resolve silently..."
if(\$IsMSI){
  \$args = "/i `"\$InstallerPath`" /qn /norestart"
  if(\$MsiProps -and \$MsiProps.Trim().Length -gt 0){ \$args += " " + \$MsiProps }
  \$proc = Start-Process -FilePath "msiexec.exe" -ArgumentList \$args -Wait -PassThru
  Write-Log "msiexec exit code: \$((\$proc.ExitCode))"
  if(\$proc.ExitCode -ne 0){ throw "MSI install returned \$((\$proc.ExitCode))" }
}else{
  \$proc = Start-Process -FilePath \$InstallerPath -ArgumentList \$ExeInstallArgs -Wait -PassThru
  Write-Log "EXE exit code: \$((\$proc.ExitCode))"
  if(\$proc.ExitCode -ne 0){ throw "EXE install returned \$((\$proc.ExitCode))" }
}

Write-Log "Installation complete."
Stop-Transcript
"@ | Set-Content -Path $InstallPs1 -Encoding UTF8 -Force

# =========================
# WRITE: Remove-ResolveOnShutdown.ps1
# =========================
@"
param(
  [string]\$LogPath = '$UninstLog',
  [string]\$ExeUninstallArgs = '$ExeUninstallArgs'
)

\$ErrorActionPreference = 'Stop'
Start-Transcript -Path \$LogPath -Append
function Write-Log(\$msg){ \$ts = (Get-Date).ToString('s'); Write-Output "[\$ts] \$msg" }

function Get-ResolveUninstallInfo {
  \$keys = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )
  foreach(\$k in \$keys){
    Get-ItemProperty -Path \$k -ErrorAction SilentlyContinue |
      Where-Object { \$_.DisplayName -match 'Resolve|GoTo Resolve|LogMeIn.*Resolve' }
  }
}

\$info = Get-ResolveUninstallInfo | Select-Object -First 1
if(-not \$info){
  Write-Log "Resolve not found; nothing to uninstall."
  Stop-Transcript
  exit 0
}

Write-Log "Found: \$((\$info.DisplayName) -replace '\s+',' ')"
if(\$info.UninstallString -match 'msiexec\.exe.*\{[0-9A-F-]+\}'){
  # Normalize to silent MSI uninstall
  if(\$info.UninstallString -match '\{[0-9A-F-]+\}'){
    \$guid = [regex]::Match(\$info.UninstallString, '\{[0-9A-F-]+\}').Value
    Write-Log "Uninstalling via MSI GUID \$guid"
    \$proc = Start-Process msiexec.exe -ArgumentList "/x \$guid /qn /norestart" -Wait -PassThru
    Write-Log "msiexec exit code: \$((\$proc.ExitCode))"
  } else {
    Write-Log "UninstallString had msiexec but no GUID; falling back to original string (may not be silent)."
    \$proc = Start-Process cmd.exe -ArgumentList "/c \$((\$info.UninstallString))" -Wait -PassThru
    Write-Log "exit code: \$((\$proc.ExitCode))"
  }
} elseif(\$info.UninstallString){
  Write-Log "Uninstalling via vendor EXE string (forcing quiet if possible)."
  # Try to append quiet flags if not present
  \$cmd = \$info.UninstallString
  if(\$cmd -notmatch '/quiet|/qn'){
    \$cmd = "\$cmd \$ExeUninstallArgs"
  }
  \$proc = Start-Process cmd.exe -ArgumentList "/c \$cmd" -Wait -PassThru
  Write-Log "exit code: \$((\$proc.ExitCode))"
} else {
  Write-Log "No uninstall string; nothing more to do."
}

Write-Log "Uninstall step complete."
Stop-Transcript
"@ | Set-Content -Path $UninstPs1 -Encoding UTF8 -Force

# =========================
# SCHEDULED TASKS
# =========================

# 1) Startup task (SYSTEM)
$startupTrigger = New-ScheduledTaskTrigger -AtStartup
$principal      = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
$actionInstall  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$InstallPs1`""
$task1          = New-ScheduledTask -Action $actionInstall -Trigger $startupTrigger -Principal $principal -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable)
Register-ScheduledTask -TaskName 'Resolve_AutoInstall_OnStartup' -InputObject $task1 -Force | Out-Null

# 2) Shutdown/uninstall trigger via ONEVENT (captures shutdown/restart user-initiated events)
# We use schtasks because New-ScheduledTaskTrigger lacks ONEVENT in PowerShell 5.x
# Triggers on System log, Provider=USER32, EventID=1074 (shutdown or restart)
$xmlQuery = "*[System[Provider[@Name='USER32'] and (EventID=1074)]]"
$cmd = 'schtasks.exe'
$args = '/Create /TN "Resolve_Uninstall_OnShutdown" /SC ONEVENT /EC System ' +
        '/MO "' + $xmlQuery + '" ' +
        '/TR "' + "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$UninstPs1`"" + '" ' +
        '/RU "SYSTEM" /RL HIGHEST /F'
Start-Process -FilePath $cmd -ArgumentList $args -Wait

Write-Host "Tasks created. Install will run at next boot; uninstall will run when the machine is shut down or restarted (User32 1074)."
