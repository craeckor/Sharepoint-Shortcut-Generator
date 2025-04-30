[CmdletBinding(SupportsShouldProcess=$true)]
Param()

$MyProgs = Get-ItemProperty 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'; $MyProgs += Get-ItemProperty 'HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
$InstalledProgs = $MyProgs.DisplayName | Sort-Object -Unique

if ($InstalledProgs -like "*PowerShell*7*") {
    Write-Host "PowerShell 7 is installed."
    pause
    exit 1
} else {
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
    $url = "https://aka.ms/powershell-release?tag=lts"
    $folderName = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $workpath = "$env:TEMP\powershell7-$folderName"
    $currentPath = $PSScriptRoot
    $downloadPath = "$workpath\powershell-7.msi"
    $curlpath = "$currentPath\curl\curl.exe"
    
    if (Test-Path $workpath) {
        Write-Host "Directory $workpath already exists. Deleting..."
        Remove-Item -Path $workpath -Recurse -Force
        New-Item -Path $workpath -ItemType Directory -Force | Out-Null
    } else {
        New-Item -Path $workpath -ItemType Directory -Force | Out-Null
    }

    Write-Host "PowerShell 7 is not installed."
    Write-Host "Installing PowerShell 7..."
    Write-Host "Please wait..."
    Write-Verbose "Extracting URL from aka.ms link..."
    $redirecturl1 = (Invoke-Expression -Command "$curlpath --connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -S -s -v `"$url`" 2>&1" | Select-String -Pattern "Location: ").Line -replace "< Location: \s*", ""
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Unable to extract URL from aka.ms link."
        pause
        exit 1
    }
    Write-Verbose "Found redirect URL: $redirecturl1"
    Write-Verbose "Extracting final download URL..."
    $redirecturl2 = (Invoke-Expression -Command "$curlpath --connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -S -s -v `"$redirecturl1`" 2>&1" | Select-String -Pattern "Location: ").Line -replace "< Location: \s*", ""
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Unable to extract final download URL."
        pause
        exit 1
    }
    Write-Verbose "Found final Github URL: $redirecturl2"
    Write-Verbose "Extract version number from URL..."
    $versionv = $redirecturl2 -split "/" | Select-Object -Last 1
    $version = $versionv -replace "v\s*", ""
    Write-Verbose "Found version number: $versionv"
    Write-Host "Downloading PowerShell 7 version $version..."
    Invoke-Expression -Command "$curlpath --connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -S -s -L -o $downloadPath `"https://github.com/PowerShell/PowerShell/releases/download/$versionv/PowerShell-$version-win-x64.msi`""
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Unable to download PowerShell 7 installer."
        pause
        exit 1
    }
    Write-Host "Download complete."
    Write-Host "Installing PowerShell 7..."
    Start-Process -FilePath msiexec.exe -ArgumentList "/i $downloadPath /passive /norestart" -Wait
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Unable to install PowerShell 7."
        pause
        exit 1
    }
    Write-Host "Installation complete."
    Write-Host "Cleaning up temporary files..."
    Remove-Item -Path $workpath -Recurse -Force
    Write-Host "Sleeping for 15 seconds..."
    Start-Sleep -Seconds 15
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
}