<#
    .SYNOPSIS
    This Script is used to provision OneDrive for Business for users in a Microsoft Entra (Azure) AD tenant.
    It handles OAuth2 authentication with Microsoft Graph API and checks if PowerShell 7 is installed. If not, it installs PowerShell 7.
    It also handles OAuth2 authentication with the Microsoft Admin Center to get access to the individual private OneDrive profiles.

    .DESCRIPTION
    This script checks if PowerShell 7 is installed, and if not, it installs it.
    It also handles OAuth2 authentication with Microsoft Graph API.
    It checks if WebView2 is installed, and if not, it installs it.
    It handles OAuth2 authentication with the Microsoft Admin Center to get access to the individual private OneDrive profiles.
    It uses the Microsoft Graph API to get all users in the tenant and allows the user to select which users to work with.
    It then provisions OneDrive for Business for the selected users.


    .PARAMETER tenantId
    Enter your Microsoft Tenant ID.
    This is the ID of your Entra Directory (Azure Active Directory) tenant.
    Docs:
    Find your tennant ID: https://learn.microsoft.com/entra/fundamentals/how-to-find-tenant

    .PARAMETER clientId
    Enter your Microsoft Client ID.
    This is the ID of your registered application in Azure AD.
    Docs:
    Get the Client ID of your registered application: https://learn.microsoft.com/azure/healthcare-apis/register-application#application-id-client-id

    .EXAMPLE
    .\OneDrive-Provisioning.ps1 -tenantId 'your-tenant-id' -clientId 'your-client-id' -Verbose

    .NOTES
    Required Scope-Permissions: files.readwrite.all, user.read.all, allsites.fullcontrol, allsites.manage, myfiles.read, myfiles.write and user.readwrite.all
    You can find these permissions in Microsoft Graph and Sharepoint API.
    files.readwrite.all and user.read.all are Microsoft Graph permissions.
    allsites.fullcontrol, allsites.manage, myfiles.read, myfiles.write and user.readwrite.all are Sharepoint API permissions.
    All Scope-Permissions require Admin Consent for the Tenant.
    It will not work without Admin Consent and if any of the permissions are missing.
    WebView2 is required for the OAuth2 authentication process.
    The redirect URI must be set to "https://login.microsoftonline.com/common/oauth2/nativeclient" in the Entra (Azure) AD application registration.

    Docs:
    Register an application in Azure AD: https://learn.microsoft.com/entra/identity-platform/quickstart-register-app
    Add a Redirect URI to an application: https://learn.microsoft.com/entra/identity-platform/how-to-add-redirect-uri
    Configure permissions for an application: https://learn.microsoft.com/entra/identity-platform/quickstart-configure-app-access-web-apis
    Find your tennant ID: https://learn.microsoft.com/entra/fundamentals/how-to-find-tenant
#> 

[CmdletBinding(SupportsShouldProcess=$true)]
Param(
    [Parameter(Mandatory=$true,
    HelpMessage = @"
Enter your Microsoft Tenant ID.
This is the ID of your Entra Directory (Azure Active Directory) tenant.
Docs:
Find your tennant ID: https://learn.microsoft.com/entra/fundamentals/how-to-find-tenant
"@)]
    [string]$tenantId,
    [Parameter(Mandatory=$true,
    HelpMessage = @"
Enter your Microsoft Client ID.
This is the ID of your registered application in Azure AD.
Docs:
Get the Client ID of your registered application: https://learn.microsoft.com/azure/healthcare-apis/register-application#application-id-client-id
"@)]
    [string]$clientId
)

$SetVerbose = $false
if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"]) {
    $SetVerbose = $true
    Write-Verbose "Verbose mode enabled"
}

function Restart-AsAdmin  {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        $argList = @(
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$PSCommandPath`"",
            "-tenantId", "`"$tenantId`"",
            "-clientId", "`"$clientId`""
        )

        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"]) {
            $argList += "-Verbose"
        }
        
        Start-Process powershell.exe -ArgumentList $argList -Verb RunAs
        exit
    }
}

function Restart-Yourself  {
    $argList = @(
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`"",
        "-tenantId", "`"$tenantId`"",
        "-clientId", "`"$clientId`""
    )

    if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"]) {
        $argList += "-Verbose"
    }
    
    Start-Process powershell.exe -ArgumentList $argList
    exit
}

Restart-AsAdmin

# Enforce TLS 1.2
# Function to configure and check the status of a TLS protocol
function Set-TLSStatus {
    param (
        [string]$Protocol,
        [string]$Type,
        [int]$EnableValue = 1
    )

    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$Protocol\$Type"

    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    Set-ItemProperty -Path $regPath -Name "Enabled" -Value $EnableValue -Force
    Set-ItemProperty -Path $regPath -Name "DisabledByDefault" -Value 0 -Force
}

# Function to check the status of a TLS protocol
function Get-TLSStatus {
    param (
        [string]$Protocol,
        [string]$Type
    )

    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$Protocol\$Type"

    if (Test-Path $regPath) {
        $enabled = Get-ItemProperty -Path $regPath -Name "Enabled" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Enabled -ErrorAction SilentlyContinue
        $disabledByDefault = Get-ItemProperty -Path $regPath -Name "DisabledByDefault" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisabledByDefault -ErrorAction SilentlyContinue

        if ($enabled -eq 1) {
            "$Protocol $Type is Enabled"
        } elseif ($disabledByDefault -eq 1) {
            "$Protocol $Type is Disabled"
        } else {
            "$Protocol $Type is Not Configured"
        }
    } else {
        "$Protocol $Type key does not exist"
    }
}

# TLS versions to configure and check
$tlsProtocols = @("TLS 1.0", "TLS 1.1", "TLS 1.2", "TLS 1.3")

# Configure TLS 1.2 and TLS 1.3
foreach ($protocol in $tlsProtocols) {
    if ($protocol -eq "TLS 1.2" -or $protocol -eq "TLS 1.3") {
        Set-TLSStatus -Protocol $protocol -Type 'Client' -EnableValue 1
        Set-TLSStatus -Protocol $protocol -Type 'Server' -EnableValue 1
    } elseif ($protocol -eq "TLS 1.0" -or $protocol -eq "TLS 1.1") {
        Set-TLSStatus -Protocol $protocol -Type 'Client' -EnableValue 0
        Set-TLSStatus -Protocol $protocol -Type 'Server' -EnableValue 0
    }
}

# Check the status for each protocol and type (Client and Server)
foreach ($protocol in $tlsProtocols) {
    Write-Output "$(Get-TLSStatus -Protocol $protocol -Type 'Client')"
    Write-Output "$(Get-TLSStatus -Protocol $protocol -Type 'Server')"
}

# .NET Framework strong cryptography setting
function Get-DotNetCryptoStatus {
    param (
        [string]$regPath
    )

    $strongCrypto = Get-ItemProperty -Path $regPath -Name "SchUseStrongCrypto" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SchUseStrongCrypto -ErrorAction SilentlyContinue
    if ($null -ne $strongCrypto) {
        if ($strongCrypto -eq 1) {
            ".NET Framework Strong Crypto at $regPath is Enabled"
        } else {
            ".NET Framework Strong Crypto at $regPath is Disabled"
        }
    } else {
        ".NET Framework Strong Crypto at $regPath is Not Configured"
    }
}

# Check .NET Framework strong cryptography settings
$netFrameworkPaths = @(
    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319"
)

foreach ($path in $netFrameworkPaths) {
    Write-Output "$(Get-DotNetCryptoStatus -regPath $path)"
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

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
    $curldefault = "--connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -S -s"
    
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
    try {
        $curlOutput = Invoke-Expression -Command "$curlpath $curldefault -v `"$url`" 2>&1"
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to extract URL from aka.ms link."
            pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        pause
        exit 1
    }
    $locationHeader = $curlOutput | Where-Object { $_ -match "Location:" } | Select-Object -First 1
    # Extract URL from location header using a more flexible pattern
    if ($locationHeader -match "Location:\s*(https?://[^\s]+)") {
        $redirecturl1 = $matches[1]
    } else {
        Write-Error "Error: Unable to extract Location header from response."
        Write-Verbose "Full response: $($curlOutput | Out-String)"
        pause
        exit 1
    }
    Write-Verbose "Found redirect URL: $redirecturl1"
    Write-Verbose "Extracting final download URL..."
    try {
        $curlOutput = Invoke-Expression -Command "$curlpath $curldefault -v `"$redirecturl1`" 2>&1"
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to extract final download URL."
            pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        pause
        exit 1
    }
    $locationHeader = $curlOutput | Where-Object { $_ -match "Location:" } | Select-Object -First 1
    # Extract URL from location header using a more flexible pattern
    if ($locationHeader -match "Location:\s*(https?://[^\s]+)") {
        $redirecturl2 = $matches[1]
    } else {
        Write-Error "Error: Unable to extract Location header from response."
        Write-Verbose "Full response: $($curlOutput | Out-String)"
        pause
        exit 1
    }
    Write-Verbose "Found final Github URL: $redirecturl2"
    Write-Verbose "Extract version number from URL..."
    $versionv = $redirecturl2 -split "/" | Select-Object -Last 1
    $version = $versionv -replace "v\s*", ""
    Write-Verbose "Found version number: $versionv"
    Write-Host "Downloading PowerShell 7 version $version..."
    try {
        Invoke-Expression -Command "$curlpath $curldefault -L -o `"$downloadPath`" `"https://github.com/PowerShell/PowerShell/releases/download/$versionv/PowerShell-$version-win-x64.msi`""
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to download PowerShell 7 installer."
            pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        pause
        exit 1
    }
    Write-Host "Download complete."
    Write-Host "Installing PowerShell 7..."
    try {
        Start-Process -FilePath msiexec.exe -ArgumentList "/i $downloadPath /passive /norestart" -Wait
        if ($LASTEXITCODE -ne 0) {
            Write-Verbose "Cleaning up temporary files..."
            Remove-Item -Path $workpath -Recurse -Force
            Write-Error "Error: Unable to install PowerShell 7."
            pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        pause
        exit 1
    }
    Write-Host "Installation complete."
    Write-Host "Cleaning up temporary files..."
    Remove-Item -Path $workpath -Recurse -Force
    Write-Host "Sleeping for 15 seconds..."
    Start-Sleep -Seconds 15
    Restart-Yourself
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
    $curldefault = "--connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -S -s"
    
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
    try {
        $curlOutput = Invoke-Expression -Command "$curlpath $curldefault -v `"$url`" 2>&1"
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to extract URL from go.microsoft.com link."
            pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        pause
        exit 1
    }
    $locationHeader = $curlOutput | Where-Object { $_ -match "Location:" } | Select-Object -First 1
    # Extract URL from location header using a more flexible pattern
    if ($locationHeader -match "Location:\s*(https?://[^\s]+)") {
        $downloadURL = $matches[1]
    } else {
        Write-Error "Error: Unable to extract Location header from response."
        Write-Verbose "Full response: $($curlOutput | Out-String)"
        pause
        exit 1
    }
    Write-Verbose "Found redirect URL: $downloadURL"
    Write-Host "Downloading WebView2 Runtime..."
    try {
        Invoke-Expression -Command "$curlpath $curldefault -L -o `"$downloadPath`" `"$downloadURL`""
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to download WebView2 Runtime installer."
            pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        pause
        exit 1
    }
    Write-Host "Download complete."
    Write-Host "Installing WebView2 Runtime..."
    Start-Process -FilePath $downloadPath -ArgumentList "/silent /install" -Wait
    if ($LASTEXITCODE -ne 0) {
        Write-Verbose "Cleaning up temporary files..."
        Remove-Item -Path $workpath -Recurse -Force
        Write-Error "Error: Unable to install WebView2 Runtime."
        pause
        exit 1
    }
    Write-Host "Installation complete."
    Write-Host "Cleaning up temporary files..."
    Remove-Item -Path $workpath -Recurse -Force
    Write-Host "Sleeping for 15 seconds..."
    Start-Sleep -Seconds 15
    Restart-Yourself
    exit 0
} else {
    Write-Host "WebView2 Runtime is already installed."
}

$workpath = $PSScriptRoot

Set-Location -Path $workpath

$curlpath = "$workpath\curl\curl.exe"
$users = @{}
$requests = @{}
$requests.succeeded = @{}
$requests.failed = @{}
$requests.counter = 0

Invoke-Expression -Command "$workpath\Import-Assemblies.ps1"
Import-Module -Name "$workpath\PSAuthClient\PSAuthClient.psd1" -Force
Import-Module -Name "$workpath\Functions\Forms-Functions.ps1" -Force
Import-Module -Name "$workpath\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.PowerShell.psd1" -Force
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Clear-WebView2Cache -Confirm:$false

$graphEndpoint = "https://graph.microsoft.com/v1.0"

function Get-OAuth2Token {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$TenantId,
        
        [Parameter(Mandatory=$true)]
        [string]$ClientId,
        
        [Parameter(Mandatory=$false)]
        [string]$RedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",
        
        [Parameter(Mandatory=$false)]
        [string]$Scope = "files.readwrite.all user.read.all allsites.fullcontrol allsites.manage myfiles.read myfiles.write user.readwrite.all",
        
        [Parameter(Mandatory=$false)]
        [switch]$UsePkce,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$CustomParameters = @{}
    )
    
    Write-Verbose "Initiating OAuth2 authentication flow"
    
    $authorization_endpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize"
    $token_endpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    
    $authParams = @{
        client_id = $ClientId
        scope = $Scope
        redirect_uri = $RedirectUri
        customParameters = $CustomParameters
    }
    
    try {
        Write-Verbose "Requesting authorization code"
        $code = Invoke-OAuth2AuthorizationEndpoint -uri $authorization_endpoint @authParams -usePkce:$UsePkce
        
        if (-not $code) {
            throw "Failed to obtain authorization code - no code was returned"
        }
        
        Write-Verbose "Authorization code obtained, exchanging for access token"
        $token = Invoke-OAuth2TokenEndpoint -uri $token_endpoint @code
        
        if (-not $token -or -not $token.access_token) {
            throw "Failed to obtain access token"
        }
        
        Write-Verbose "Access token successfully obtained."
        
        # Return a custom object with token information and expiration details
        $tokenExpiresAt = (Get-Date).AddSeconds($token.expires_in)
        
        return [PSCustomObject]@{
            AccessToken = $token.access_token
            TokenType = $token.token_type
            ExpiresIn = $token.expires_in
            ExpiresAt = $tokenExpiresAt
            RefreshToken = $token.refresh_token
            Scope = $token.scope
            IdToken = $token.id_token
        }
    }
    catch {
        Write-Error "OAuth2 authentication failed: $($_.Exception.Message)"
        Write-Verbose "Full error details: $($_)"
        
        if ($_.Exception.Message -match "AADSTS65001") {
            Write-Warning "Consent required - ensure all required permissions have admin consent in Azure AD"
        }
        elseif ($_.Exception.Message -match "AADSTS700016") {
            Write-Warning "Application not found in directory - verify the Client ID is correct"
        }
        elseif ($_.Exception.Message -match "AADSTS90002") {
            Write-Warning "Tenant not found - verify the Tenant ID is correct"
        }
        
        throw $_
    }
}

function Send-GraphRequest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Uri,
        [Parameter(Mandatory=$true)]
        [ValidateSet("GET", "POST", "PATCH")]
        [string]$Method,
        [Parameter(Mandatory=$false)]
        [string]$AccessToken,
        [Parameter(Mandatory=$false)]
        [string]$Body,
        [Parameter(Mandatory=$false)]
        [string]$Cookie
    )
    
    $maxRetries = 5
    $retryCount = 0
    $success = $false
    $tokenRefreshed = $false
    
    # Increment request counter before making the request
    $script:requests.counter++
    $currentRequestId = $script:requests.counter
    
    do {
        try {
            if ($retryCount -gt 0) {
                Write-Verbose "Retry attempt $retryCount of $maxRetries for request #$currentRequestId to: $Uri"
                # Add exponential backoff with jitter
                $backoffSeconds = [math]::Pow(2, $retryCount) + (Get-Random -Minimum 0 -Maximum 1000) / 1000
                Write-Verbose "Waiting $backoffSeconds seconds before retry..."
                Start-Sleep -Seconds $backoffSeconds
            } else {
                Write-Verbose "Making request #$currentRequestId to: $Uri"
            }

            # For POST and PATCH methods, use Invoke-RestMethod if no cookie is needed
            if (($Method -eq "POST" -or $Method -eq "PATCH") -and -not $Cookie) {
                $headers = @{
                    "Content-Type" = "application/json"
                    "Accept" = "application/json"
                }

                # Add Authorization header if AccessToken is provided
                if ($AccessToken) {
                    $headers["Authorization"] = "Bearer $AccessToken"
                }

                $params = @{
                    Method = $Method
                    Uri = $Uri
                    Headers = $headers
                    ContentType = "application/json"
                }

                # Add Body if provided
                if ($Body) {
                    $params.Body = $Body
                }

                try {
                    $response = Invoke-RestMethod @params -ErrorAction Stop

                    # Simulate success since Invoke-RestMethod throws on non-success codes
                    $httpCode = "200"
                    # Convert response to JSON string to maintain consistency with curl responses
                    $responseBody = $response | ConvertTo-Json -Depth 10
                } 
                catch [System.Net.WebException] {
                    $responseStream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($responseStream)
                    $responseBody = $reader.ReadToEnd()
                    $httpCode = [int]$_.Exception.Response.StatusCode
                }
            } else {
                # Use curl for GET requests or when cookies are needed
                # Build basic curl cmd
                $curlCmd = "$curlpath --connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -o - -S -s -w '%{http_code}' -X $Method"

                # Add Authorization header if AccessToken is provided
                if ($AccessToken) {
                    $curlCmd += " -H `"Authorization: Bearer $AccessToken`""
                }

                # Add Body if provided
                if ($Body) {
                    $curlCmd += " -d `"$Body`""
                }

                # Add Cookie if provided
                if ($Cookie) {
                    $curlCmd += " -b `"$Cookie`""
                }

                # Add standard headers
                $curlCmd += " -H `"Content-Type: application/json`" -H `"Accept: application/json`""

                # Add the URI
                $curlCmd += " `"$Uri`""
                
                $rawResponse = Invoke-Expression -Command "$curlCmd"
                
                if ($LASTEXITCODE -ne 0) {
                    Write-Verbose "Curl command failed with exit code $LASTEXITCODE"
                    throw "Curl command failed with exit code $LASTEXITCODE"
                }
                
                # Extract the HTTP status code (last 3 characters)
                $httpCode = $rawResponse.Substring($rawResponse.Length - 3)
                
                # Extract the actual response body (everything except the last 3 characters)
                $responseBody = $rawResponse.Substring(0, $rawResponse.Length - 3)
            }
            
            # Check if HTTP status code is 403 - try refreshing token once
            if ($httpCode -eq "403" -and -not $tokenRefreshed -and -not $Cookie) {
                Write-Verbose "HTTP request #$currentRequestId failed with status code: 403 (Forbidden) - Attempting to refresh token"
                
                try {
                    # Request a new token
                    $newTokenInfo = Get-OAuth2Token -TenantId $tenantId -ClientId $clientId -Verbose:$SetVerbose
                    $AccessToken = $newTokenInfo.AccessToken
                    $tokenRefreshed = $true
                    Write-Verbose "Token refreshed successfully, will retry request with new token"
                    
                    # Reset retry count to give full retry attempts with new token
                    $retryCount = 0
                    continue
                }
                catch {
                    Write-Verbose "Failed to refresh token: $($_.Exception.Message)"
                    
                    # Store failed request information
                    $script:requests.failed[$currentRequestId] = @{
                        Uri = $Uri
                        Timestamp = Get-Date
                        Error = "HTTP request failed with status code: 403 (Forbidden) - Token refresh failed: $($_.Exception.Message)"
                        HttpCode = $httpCode
                        Response = $responseBody
                    }
                    
                    throw "Access denied (403 Forbidden). This may indicate missing permissions or admin consent. Please ensure your application has the required permissions: files.readwrite.all, user.read.all, allsites.fullcontrol, allsites.manage, myfiles.read, myfiles.write, user.readwrite.all."
                }
            }
            # Handle 403 error after token refresh attempt
            elseif ($httpCode -eq "403" -and $tokenRefreshed -and -not $Cookie) {
                # Store failed request information
                $script:requests.failed[$currentRequestId] = @{
                    Uri = $Uri
                    Timestamp = Get-Date
                    Error = "HTTP request failed with status code: 403 (Forbidden) - Even after token refresh"
                    HttpCode = $httpCode
                    Response = $responseBody
                }
                
                Write-Verbose "HTTP request #$currentRequestId failed with status code: 403 (Forbidden) - Even after token refresh"
                throw "Access denied (403 Forbidden) even after token refresh. This indicates insufficient permissions for this operation. Please verify that all required permissions have been granted and admin consent has been provided in Azure AD."
            }
            # Special handling for 409 (Conflict) errors - track as error but return the response
            elseif ($httpCode -eq "409") {
                Write-Verbose "HTTP request #$currentRequestId returned status code: 409 (Conflict) - Will track as error but return the response"
                
                # Store in failed requests but don't throw an error
                $script:requests.failed[$currentRequestId] = @{
                    Uri = $Uri
                    Timestamp = Get-Date
                    Error = "HTTP request returned status code: 409 (Conflict)"
                    HttpCode = $httpCode
                    Response = $responseBody
                }
                
                # Set success to true so we exit the retry loop
                $success = $true
                
                # Return the response body despite the error
                return $responseBody
            }
            # Check if HTTP status code indicates other errors that we should retry
            elseif (-not ($httpCode -match "^(2\d\d)$")) {
                Write-Verbose "HTTP request #$currentRequestId failed with status code: $httpCode - Will retry"
                throw "HTTP request failed with status code: $httpCode"
            }
            
            # If we got here, the request was successful
            $success = $true
            
            # Store successful request information
            $script:requests.succeeded[$currentRequestId] = @{
                Uri = $Uri
                Timestamp = Get-Date
                HttpCode = $httpCode
                Response = $responseBody
                RetryCount = $retryCount
                TokenRefreshed = $tokenRefreshed
            }
            
            Write-Verbose "Request #$currentRequestId completed successfully with status code: $httpCode after $retryCount retries"
            
            # Return just the response body
            return $responseBody
        }
        catch {
            $retryCount++
            
            # Check if we've reached max retries or if it's a terminal 403 error (after token refresh)
            if ($retryCount -gt $maxRetries -or ($_.Exception.Message -match "403 \(Forbidden\)" -and $tokenRefreshed)) {
                # Store failed request information if it's the final attempt
                if (-not $script:requests.failed.ContainsKey($currentRequestId)) {
                    $script:requests.failed[$currentRequestId] = @{
                        Uri = $Uri
                        Timestamp = Get-Date
                        Error = $_.Exception.Message
                        HttpCode = if ($httpCode) { $httpCode } else { $null }
                        Response = if ($responseBody) { $responseBody } else { $null }
                        RetryCount = $retryCount - 1
                        TokenRefreshed = $tokenRefreshed
                    }
                }
                
                Write-Error "Error in GraphAPI request after $($retryCount-1) retries: $($_.Exception.Message)"
                
                # Instead of returning the error message as if it were a valid response,
                # throw the error so the caller can handle it properly
                throw "Request failed after $maxRetries retries: $($_.Exception.Message)"
            }
            
            Write-Verbose "Request failed. Error: $($_.Exception.Message). Retrying..."
        }
    } while (-not $success -and $retryCount -le $maxRetries)
}

try {
    Write-Host "Starting OAuth2 authentication process..."
    $tokenInfo = Get-OAuth2Token -TenantId $tenantId -ClientId $clientId -Verbose:$SetVerbose
    $accessToken = $tokenInfo.AccessToken
    Write-Verbose "Token will expire at: $($tokenInfo.ExpiresAt)"
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

Write-Host "Getting all users in tenant..."
try {
    # Get all users in the tenant
    $allUsers = Send-GraphRequest -Method GET -Uri "$graphEndpoint/users" -AccessToken "$accessToken" -Verbose:$SetVerbose
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to get users from Graph API."
        pause
        exit 1
    }
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

# Parse the JSON response
$allUsersObj = $allUsers | ConvertFrom-Json

# Generate and display the table
$usersAllTable = ($allUsersObj.value | Select-Object DisplayName, mail, id | Format-Table -AutoSize | Out-String).Trim()
Write-Verbose "All users in tenant:`n$usersAllTable"

Write-Host "Showing user selection form..."
Write-Host "Please select the users you want to work with..."

try {
    # Only call ONCE and assign result
    $selectedUsers = Show-UserSelectionForm -userList $allUsersObj.value -Verbose:$SetVerbose
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to show user selection form."
        pause
        exit 1
    }
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

if (-not $selectedUsers -or $selectedUsers.Count -eq 0) {
    Write-Error "No users selected. Exiting..."
    pause
    exit
}

$users.list = $selectedUsers

# Fix for Write-Verbose formatting issues
Write-Verbose "Users object: $($users | ConvertTo-Json -Depth 10 -Compress)"
Write-Verbose "Selected users count: $($users.list.Count)"
# Format as string first, then output with Write-Verbose
$usersTable = ($users.list | Select-Object DisplayName, Mail, Id | Format-Table -AutoSize | Out-String).Trim()
Write-Verbose "Selected users:`n$usersTable"

Write-Host "Get root site name..."
# Get root site name
try {
    $rootSite = Send-GraphRequest -Method GET -Uri "$graphEndpoint/sites/root" -AccessToken "$accessToken" -Verbose:$SetVerbose
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to get root site from Graph API."
        pause
        exit 1
    }
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

$rootSiteObj = $rootSite | ConvertFrom-Json
$rootSiteName = $rootSiteObj.siteCollection.hostname
Write-Verbose  "Root site name: $rootSiteName"

# Create Admin SharePoint site URL
$tenantName = $rootSiteName.Split('.')[0]
$adminSiteUrl = "$tenantName-admin.sharepoint.com"
Write-Verbose "Admin SharePoint site URL: $adminSiteUrl"
Write-Host "Admin SharePoint site URL: $adminSiteUrl" -ForegroundColor Cyan

# Connect to SPOService
try {
    Write-Host "Connecting to SharePoint Online Admin Center..."
    Connect-SPOService -Url "https://$($adminSiteUrl)" -ErrorAction Stop
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to connect to SharePoint Online Admin Center."
        pause
        exit 1
    }
    Write-Verbose "Connected to SharePoint Online Admin Center successfully."
} catch {
    Write-Error "Error: Unable to connect to SharePoint Online Admin Center. $($_.Exception.Message)"
    pause
    exit 1
}

$userlists = @{}

# Create user lists with max 199 users per list (Graph API pagination handling)
if ($users.list.Count -gt 0) {
    Write-Host "Creating user lists for processing..."
    $batchSize = 199
    $batchCount = [Math]::Ceiling($users.list.Count / $batchSize)
    
    # Create an array to hold all email addresses
    $allEmails = @()
    
    for ($i = 0; $i -lt $batchCount; $i++) {
        $startIndex = $i * $batchSize
        $endIndex = [Math]::Min(($i + 1) * $batchSize - 1, $users.list.Count - 1)
        $currentBatchSize = $endIndex - $startIndex + 1
        
        $userlists["batch_$i"] = @()
        
        for ($j = $startIndex; $j -le $endIndex; $j++) {
            # Extract ONLY the Mail property from each user
            $email = $users.list[$j].Mail
            $userlists["batch_$i"] += $email
            $allEmails += $email
        }
        
        Write-Verbose "Created batch_$i with $currentBatchSize emails"
    }
    
    Write-Host "Created $batchCount email batch(es) for processing" -ForegroundColor Green
    Write-Verbose "Total unique email addresses: $($allEmails.Count)"
    Write-Verbose "First 5 email addresses: $($allEmails | Select-Object -First 5 -Unique)"
} else {
    Write-Host "No users selected, nothing to process" -ForegroundColor Yellow
}

# Provision OneDrive personal sites for each batch of users
if ($userlists.Count -gt 0) {
    Write-Host "`nProvisioning OneDrive personal sites for all users..." -ForegroundColor Cyan
    
    $finalUserMailList = @()
    $totalProcessed = 0
    $batchProcessed = 0
    
    foreach ($batchKey in $userlists.Keys) {
        $list = $userlists[$batchKey]
        $batchSize = $list.Count
        $totalProcessed += $batchSize
        $batchProcessed++
        
        Write-Host "Processing batch $($batchKey): $batchSize users (batch $batchProcessed of $($userlists.Count))" -ForegroundColor Yellow
        
        try {
            # Add emails to the final tracking list
            $finalUserMailList += $list
            
            # Request OneDrive site provisioning with -NoWait parameter for efficiency
            Request-SPOPersonalSite -UserEmails $list -NoWait
            if ($LASTEXITCODE -ne 0) {
                Write-Error "Error requesting OneDrive provisioning for batch $($batchKey)."
                Pause
                exit 1
            }
            
            Write-Host "Successfully requested OneDrive provisioning for batch $batchKey" -ForegroundColor Green
            Write-Verbose "Emails in this batch: $($list -join ', ')"
        }
        catch {
            Write-Error "Error requesting OneDrive provisioning for batch $($batchKey): $($_.Exception.Message)"
            Pause
            exit 1
        }
    }
    
    Write-Host "`nOneDrive provisioning summary:" -ForegroundColor Cyan
    Write-Host "- Total users processed: $totalProcessed" -ForegroundColor White
    Write-Host "- Total batches processed: $batchProcessed" -ForegroundColor White
    Write-Host "- Users will receive their OneDrive sites asynchronously" -ForegroundColor White
    
    Write-Verbose "OneDrive provisioning requested for the following emails: $($finalUserMailList -join ', ')"
    Write-Host "`nNote: This might take some time to complete. Please check the Admin Center for status updates."  -ForegroundColor Yellow
}
else {
    Write-Host "`nNo user batches available for OneDrive provisioning." -ForegroundColor Yellow
}

# Display request statistics in a more appropriate way
Write-Host "Request Statistics:" -ForegroundColor Cyan
Write-Host "Total requests made: $($requests.counter)" -ForegroundColor Cyan

# Only display succeeded requests if there are any
if ($requests.succeeded.Count -gt 0) {
    Write-Host "`nSuccessful Requests:" -ForegroundColor Green
    $successTable = $requests.succeeded.GetEnumerator() | 
        Select-Object @{Name='RequestId';Expression={$_.Key}}, 
                     @{Name='Uri';Expression={$_.Value.Uri}}, 
                     @{Name='Timestamp';Expression={$_.Value.Timestamp}}, 
                     @{Name='HttpCode';Expression={$_.Value.HttpCode}} |
        Format-Table -AutoSize | Out-String
    Write-Host $successTable
}
else {
    Write-Host "`nNo successful requests recorded." -ForegroundColor Yellow
}

# Only display failed requests if there are any
if ($requests.failed.Count -gt 0) {
    Write-Host "`nFailed Requests:" -ForegroundColor Red
    $failedTable = $requests.failed.GetEnumerator() | 
        Select-Object @{Name='RequestId';Expression={$_.Key}}, 
                     @{Name='Uri';Expression={$_.Value.Uri}}, 
                     @{Name='Timestamp';Expression={$_.Value.Timestamp}}, 
                     @{Name='Error';Expression={$_.Value.Error}}, 
                     @{Name='HttpCode';Expression={$_.Value.HttpCode}} |
        Format-Table -AutoSize | Out-String
    Write-Host $failedTable
}
else {
    Write-Host "`nNo failed requests recorded." -ForegroundColor Yellow
}

# Output summary
Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "- Total Requests: $($requests.counter)" -ForegroundColor White
Write-Host "- Successful: $($requests.succeeded.Count)" -ForegroundColor Green
Write-Host "- Failed: $($requests.failed.Count)" -ForegroundColor Red

pause