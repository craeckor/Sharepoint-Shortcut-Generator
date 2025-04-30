<#
    .SYNOPSIS
    This Script is ablel to handle adding Shortcuts to private Tennant OneDrive profiles redirected to the Sharepoint Online files. It supports multiple Users and multiple Sharepoint Sites with multiple Sharepoint Directories.
    It handles OAuth2 authentication with Microsoft Graph API and checks if PowerShell 7 is installed. If not, it installs PowerShell 7.
    It also handles OAuth2 authentication with the Microsoft Admin Center to get access to the individual private OneDrive profiles.

    .DESCRIPTION
    This script checks if PowerShell 7 is installed, and if not, it installs it.
    It also handles OAuth2 authentication with Microsoft Graph API.

    .PARAMETER tenantId
    Enter your Microsoft Tenant ID.

    .PARAMETER clientId
    Enter your Microsoft Client ID.

    .EXAMPLE
    .\oauth.ps1 -tenantId 'your-tenant-id' -clientId 'your-client-id'

    .NOTES
    Required Scope-Permissions: files.readwrite.all  user.read allsites.fullcontrol myfiles.read myfiles.write user.readwrite.all
    All Scope-Permissions require Admin Consent for the Tenant.
    WebView2 is required for the OAuth2 authentication process.
#> 

[CmdletBinding(SupportsShouldProcess=$true)]
Param(
    [Parameter(Mandatory=$true,
    HelpMessage = "Enter your Microsoft Tennant ID.")]
    [string]$tenantId,
    [Parameter(Mandatory=$true,
    HelpMessage = "Enter your Microsoft Client ID.")]
    [string]$clientId
)

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

$MyProgs = Get-ItemProperty 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'; $MyProgs += Get-ItemProperty 'HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
$InstalledProgs = $MyProgs.DisplayName | Sort-Object -Unique

$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

function Get-WebView2Installation {
    [CmdletBinding()]
    param ()

    $runtimeNeeded = $true
    $verifyRuntime = $false
    $version = $null

    # See: https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/distribution#detect-if-a-suitable-webview2-runtime-is-already-installed

    if ([Environment]::Is64BitOperatingSystem) {
        # Test x64
        $regPath = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
        try {
            $version = (Get-ItemProperty -Path $regPath -ErrorAction Stop).pv
            $verifyRuntime = $true
        } catch {}
    }

    if (-not $verifyRuntime) {
        # Test x32
        $regPath = 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
        try {
            $version = (Get-ItemProperty -Path $regPath -ErrorAction Stop).pv
            $verifyRuntime = $true
        } catch {}
    }

    if ($verifyRuntime) {
        if ($version -and $version -ne '0.0.0.0') {
            Write-Verbose "WebView2 Runtime is installed (version: $version)"
            $runtimeNeeded = $false
        } else {
            Write-Verbose "WebView2 Runtime needs to be downloaded and installed"
        }
    }

    return $runtimeNeeded
}

if ($InstalledProgs -like "*PowerShell*7*") {
    Write-Host "PowerShell 7 is installed."
} else {
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
        Write-Debug "Cleaning up temporary files..."
        Remove-Item -Path $workpath -Recurse -Force
        pause
        exit 1
    }
    Write-Host "Installation complete."
    Write-Host "Cleaning up temporary files..."
    Remove-Item -Path $workpath -Recurse -Force
    Write-Host "Sleeping for 15 seconds..."
    Start-Sleep -Seconds 15
    Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit 0
}

# Check if WebView2 is needed
$webview2Needed = Get-WebView2Installation
if ($webview2Needed) {
    $url = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
    $folderName = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $workpath = "$env:TEMP\webview2-$folderName"
    $currentPath = $PSScriptRoot
    $downloadPath = "$workpath\MicrosoftEdgeWebview2Setup.exe"
    $curlpath = "$currentPath\curl\curl.exe"
    
    if (Test-Path $workpath) {
        Write-Host "Directory $workpath already exists. Deleting..."
        Remove-Item -Path $workpath -Recurse -Force
        New-Item -Path $workpath -ItemType Directory -Force | Out-Null
    } else {
        New-Item -Path $workpath -ItemType Directory -Force | Out-Null
    }

    Write-Host "WebView2 Runtime is not installed."
    Write-Host "Installing WebView2 Runtime..."
    Write-Host "Please wait..."
    Write-Verbose "Extracting URL from go.microsoft.com link..."
    $downloadURL = (Invoke-Expression -Command "$curlpath --connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -S -s -v `"$url`" 2>&1" | Select-String -Pattern "Location: ").Line -replace "< Location: \s*", ""
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Unable to extract URL from go.microsoft.com link."
        pause
        exit 1
    }
    Write-Verbose "Found redirect URL: $downloadURL"
    Write-Host "Downloading WebView2 Runtime..."
    Invoke-Expression -Command "$curlpath --connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -S -s -L -o $downloadPath `"$downloadURL`""
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Unable to download WebView2 Runtime installer."
        pause
        exit 1
    }
    Write-Host "Download complete."
    Write-Host "Installing WebView2 Runtime..."
    Start-Process -FilePath $downloadPath -ArgumentList "/silent /install" -Wait
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Unable to install WebView2 Runtime."
        Write-Debug "Cleaning up temporary files..."
        Remove-Item -Path $workpath -Recurse -Force
        pause
        exit 1
    }
    Write-Host "Installation complete."
    Write-Host "Cleaning up temporary files..."
    Remove-Item -Path $workpath -Recurse -Force
    Write-Host "Sleeping for 15 seconds..."
    Start-Sleep -Seconds 15
    Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit 0
} else {
    Write-Host "WebView2 Runtime is already installed."
}

$workpath = $PSScriptRoot

Set-Location -Path $workpath

$curlpath = "$workpath\curl\curl.exe"

Invoke-Expression -Command "$workpath\import-assemblies.ps1"
Import-Module -Name $workpath\PSAuthClient\PSAuthClient.psd1 -Force

Clear-WebView2Cache -Confirm:$false

$authorization_endpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize"
$token_endpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

$splat = @{
    client_id = "$clientId"
    scope = "files.readwrite.all  user.read allsites.fullcontrol myfiles.read myfiles.write user.readwrite.all"
    redirect_uri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
    customParameters = @{}
}

try {
    $code = Invoke-OAuth2AuthorizationEndpoint -uri $authorization_endpoint @splat -usePkce:$false
    $token = Invoke-OAuth2TokenEndpoint -uri $token_endpoint @code
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    pause
    exit 1
}

Write-Host "Access Token: $($token.access_token)"

pause