<#
    .SYNOPSIS
    This script adds shortcuts to private OneDrive profiles in your tenant, redirecting them to SharePoint Online folders. It supports multiple users, multiple SharePoint sites, and multiple SharePoint folders.
    The script handles OAuth2 authentication with the Microsoft Graph API and checks if PowerShell 7 is installed, installing it if necessary.
    It also manages OAuth2 authentication with the Microsoft Admin Center to access individual private OneDrive profiles.

    .DESCRIPTION
    This script checks for the presence of PowerShell 7 and installs it if missing.
    It manages OAuth2 authentication with the Microsoft Graph API.
    The script checks for the WebView2 runtime and installs it if required.
    It authenticates with the Microsoft Admin Center to access private OneDrive profiles.
    Using the Microsoft Graph API, it retrieves all users in the tenant and allows you to select which users to process.
    It allows you to select which SharePoint sites and folders to work with.
    The script then generates shortcuts for the selected users, sites, and folders.

    .PARAMETER tenantId
    Your Microsoft Tenant ID.
    This is the ID of your Entra Directory (Azure Active Directory) tenant.
    Documentation:
    How to find your tenant ID: https://learn.microsoft.com/entra/fundamentals/how-to-find-tenant

    .PARAMETER clientId
    Your Microsoft Client ID.
    This is the ID of your registered application in Azure AD.
    Documentation:
    How to get the Client ID of your registered application: https://learn.microsoft.com/azure/healthcare-apis/register-application#application-id-client-id

    .EXAMPLE
    .\Sharepoint-Shortcut-Generator.ps1 -tenantId 'your-tenant-id' -clientId 'your-client-id' -Verbose

    .NOTES
    Required permissions (scopes): files.readwrite.all, user.read.all, allsites.fullcontrol, allsites.manage, myfiles.read, myfiles.write, and user.readwrite.all.
    These permissions are found in the Microsoft Graph and SharePoint APIs.
    - files.readwrite.all and user.read.all are Microsoft Graph permissions.
    - allsites.fullcontrol, allsites.manage, myfiles.read, myfiles.write, and user.readwrite.all are SharePoint API permissions.
    All scope permissions require admin consent for the tenant.
    The script will not work without admin consent or if any required permissions are missing.
    WebView2 is required for the OAuth2 authentication process.
    The redirect URI must be set to "https://login.microsoftonline.com/common/oauth2/nativeclient" in the Entra (Azure) AD application registration.

    Documentation:
    Register an application in Azure AD: https://learn.microsoft.com/entra/identity-platform/quickstart-register-app
    Add a redirect URI to an application: https://learn.microsoft.com/entra/identity-platform/how-to-add-redirect-uri
    Configure permissions for an application: https://learn.microsoft.com/entra/identity-platform/quickstart-configure-app-access-web-apis
    How to find your tenant ID: https://learn.microsoft.com/entra/fundamentals/how-to-find-tenant
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
$sites = @{}
$requests = @{}
$requests.succeeded = @{}
$requests.failed = @{}
$requests.counter = 0

Invoke-Expression -Command "$workpath\Import-Assemblies.ps1"
Import-Module -Name "$workpath\PSAuthClient\PSAuthClient.psd1" -Force
Import-Module -Name "$workpath\Functions\Forms-Functions.ps1" -Force
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

Write-Host "Getting all sites in tenant..."

try {
    $allSites = Send-GraphRequest -Method GET -Uri "$graphEndpoint/sites?search=*" -AccessToken "$accessToken" -Verbose:$SetVerbose
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to get sites from Graph API."
        pause
        exit 1
    }
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

$allSitesObj = $allSites | ConvertFrom-Json

# Parse the structured site IDs into their components
foreach ($site in $allSitesObj.value) {
    if ($site.id -match "([^,]+),([^,]+),(.+)") {
        $site | Add-Member -NotePropertyName "domainId" -NotePropertyValue $Matches[1] -Force
        $site | Add-Member -NotePropertyName "siteId" -NotePropertyValue $Matches[2] -Force
        $site | Add-Member -NotePropertyName "webId" -NotePropertyValue $Matches[3] -Force
    }
}

$sitesAllTable = ($allSitesObj.value | Select-Object DisplayName, WebUrl | Format-Table -AutoSize | Out-String).Trim()
Write-Verbose "All sites in tenant:`n$sitesAllTable"

Write-Host "Showing site selection form..."
Write-Host "Please select the sites you want to work with..."

try {
    $selectedSites = Show-SiteSelectionForm -SiteList $allSitesObj.value -Verbose:$SetVerbose
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to show site selection form."
        pause
        exit 1
    }
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

if (-not $selectedSites -or $selectedSites.Count -eq 0) {
    Write-Error "No sites selected. Exiting..."
    pause
    exit
}

$sites.list = $selectedSites

# Make sure selected sites include the ID components
foreach ($site in $sites.list) {
    if (-not ($site.PSObject.Properties.Name -contains "domainId") -and 
        $site.id -match "([^,]+),([^,]+),(.+)") {
        $site | Add-Member -NotePropertyName "domainId" -NotePropertyValue $Matches[1] -Force
        $site | Add-Member -NotePropertyName "siteId" -NotePropertyValue $Matches[2] -Force
        $site | Add-Member -NotePropertyName "webId" -NotePropertyValue $Matches[3] -Force
    }
}

# Format as string first, then output with Write-Verbose
$sitesTable = ($sites.list | Format-Table -Property DisplayName, WebUrl, id -AutoSize | Out-String).Trim()
Write-Verbose "Selected sites:`n$sitesTable"

Write-Host "Getting drives for selected sites..."

# Create a new array to store sites with available document libraries
$validSites = @()

foreach ($site in $sites.list) {
    $site | Add-Member -NotePropertyName "DriveId" -NotePropertyValue $null -Force

    try{
        $drive = Send-GraphRequest -Method GET -Uri "$graphEndpoint/sites/$($site.id)/drives" -AccessToken "$accessToken" -Verbose:$SetVerbose
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to get drives from Graph API."
            Write-Verbose "Error details: $drive"
            Pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        Pause
        exit 1
    }

    $driveObj = $drive | ConvertFrom-Json

    # Check if drive includes a value where name is Documents
    $documentsLibrary = $driveObj.value | Where-Object {$_.name -eq "Documents"}
    
    if ($documentsLibrary) {
        $site.DriveId = $documentsLibrary.id
        Write-Verbose "Documents library found for site: $($site.DisplayName)"
        $validSites += $site
    } else {
        Write-Warning "No Documents library found for site: $($site.DisplayName) - Removing from site list"
    }
}

# Replace the original sites list with only the valid sites
$sites.list = $validSites

Write-Verbose "Retained $(($sites.list | Measure-Object).Count) sites with available Documents libraries"

# Format as string first, then output with Write-Verbose
$sitesTable = ($sites.list | Format-Table -Property DisplayName, DriveId, WebUrl, id -AutoSize | Out-String).Trim()
Write-Verbose "Selected sites:`n$sitesTable"

Write-Host "Getting permissions for selected sites..."

foreach ($site in $sites.list) {
    $site | Add-Member -NotePropertyName "Permissions" -NotePropertyValue @{} -Force

    Write-Verbose "Getting permissions for site: $($site.displayName)"
    
    # Get User Information List
    try {
        $permissionslist = Send-GraphRequest -Method GET -Uri "$(([uri]::new("$graphEndpoint/sites/$($site.id)/lists?filter=displayName eq 'User Information List'")).AbsoluteUri)" -AccessToken "$accessToken" -Verbose:$SetVerbose
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to get permissions list from Graph API."
            pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        Pause
        exit 1
    }
    $permissionslistObj = $permissionslist | ConvertFrom-Json
    $permissionslistId = $permissionslistObj.value | Where-Object {$_.displayName -eq "User Information List"} | Select-Object -ExpandProperty id
    
    if (-not $permissionslistId) {
        Write-Error "Error: No 'User Information List' found for site: $($site.displayName)"
        pause
        exit 1
    }
    
    # Get list items with all fields
    try {
        $permissions = Send-GraphRequest -Method GET -Uri "$graphEndpoint/sites/$($site.id)/lists/$permissionslistId/items?expand=fields" -AccessToken "$accessToken" -Verbose:$SetVerbose
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to get permissions from Graph API."
            pause
            exit 1
        }
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        Pause
        exit 1
    }
    $permissionsObj = $permissions | ConvertFrom-Json
    
    # Extract only the fields property from each item
    $fieldsOnly = $permissionsObj.value | Select-Object -ExpandProperty fields
    
    # Store in the site object
    $site.Permissions = $fieldsOnly
    
    # Display first 5 items with their key properties
    Write-Verbose "Fields for first 5 users in site '$($site.displayName)':"
    $fieldsTable = ($fieldsOnly | Select-Object | 
        Select-Object EMail, Name, FirstName, LastName -First 5 |
        Format-Table -AutoSize | Out-String).Trim()
    Write-Verbose $fieldsTable
    
    # Show total count
    Write-Verbose "Total users in permissions list: $($fieldsOnly.Count)"
}

Write-Host "Associating users with sites..."

$totalaccessCount = 0

foreach ($user in $users.list) {
    $user | Add-Member -NotePropertyName "sites" -NotePropertyValue @{} -Force
    $user | Add-Member -NotePropertyName "hasAccess" -NotePropertyValue $null -Force
    $userEmail = $user.Mail
    Write-Verbose "Looking for sites with access for user: $($user.displayName) ($userEmail)"

    foreach ($site in $sites.list) {
        # Find matching user in the site permissions by email
        $matchingPermission = $site.Permissions | Where-Object { $_.EMail -eq $userEmail }
        
        if ($matchingPermission) {
            # Add site to user's sites collection - simpler structure with just the essentials
            $user.sites[$site.id] = @{
                siteId = $site.id
                siteName = $site.displayName
                driveId = $site.DriveId
            }
            
            Write-Verbose "  Added site: $($site.displayName) to user $($user.displayName)'s access list"
        } else {
            Write-Verbose "  User $($user.displayName) does not have access to site: $($site.displayName)"
        }
    }
    
    # Show summary of sites for this user
    $siteCount = $user.sites.Count
    if ($siteCount -gt 0) {
        $user.hasAccess = $true
        Write-Verbose "User $($user.displayName) has access to $siteCount sites"
        $totalaccessCount++
    } else {
        $user.hasAccess = $false
        Write-Warning "User $($user.displayName) does not have access to any of the selected sites"
    }
}

if ($totalaccessCount -eq 0) {
    Write-Error "No users have access to any of the selected sites. Exiting..."
    pause
    exit
}

# Get DriveList with proper structure for the folders form
$driveList = @()
foreach ($site in $sites.list) {
    if ($site.DriveId) {
        $driveList += [PSCustomObject]@{
            id = $site.DriveId
            name = "Documents"
            siteDisplayName = $site.DisplayName
            siteId = $site.id
            webUrl = $site.WebUrl
            domainId = $site.domainId
            webId = $site.webId
            domainSiteId = $site.siteId
        }
    }
}

Write-Host "Getting folders for selected drives..."
Write-Host "Please select the folders you want to work with..."
try {
    # Show the folder selection form
    $selectedFolders = Show-FolderSelectionForm -DriveList $driveList -AccessToken $accessToken -Verbose:$SetVerbose
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to get folders from Graph API."
        pause
        exit 1
    }
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

# Display selected folders
if (-not $selectedFolders -or $selectedFolders.Count -eq 0) {
    Write-Error "No folders selected. Exiting..."
    pause
    exit
}

# Print selected folders
Write-Host "Selected Folders:" -ForegroundColor Cyan
$foldersTable = $selectedFolders | Format-Table -Property FolderName,SiteDisplayName -AutoSize | Out-String
Write-Host $foldersTable

Write-Host "Processing selected folders..."

foreach ($folder in $selectedFolders) {
    try {
        Write-Verbose "Folder WebUrl: $($folder.webUrl)"
        Write-Verbose "SharePointId: $($folder.siteId)"

        # Get Documents list ID
        $documentsList = Send-GraphRequest -Method GET -Uri "$(([uri]::new("$graphEndpoint/sites/$($folder.siteId)/lists?filter=displayName eq 'Documents'")).AbsoluteUri)" -AccessToken "$accessToken" -Verbose:$SetVerbose
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to get Documents list from Graph API."
            Write-Verbose "Error details: $documentsList"
            Pause
            exit 1
        }
        $documentsListObj = $documentsList | ConvertFrom-Json
        $documentsListId = $documentsListObj.value | Where-Object {$_.name -eq "Shared Documents"} | Select-Object -ExpandProperty id

        # Get eTag and webUrl from Shared Documents list
        $documentsListItems = Send-GraphRequest -Method GET -Uri "$graphEndpoint/sites/$($folder.siteId)/lists/$documentsListId/items?select=eTag,webUrl" -AccessToken "$accessToken" -Verbose:$SetVerbose
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Error: Unable to get Documents list items from Graph API."
            Write-Verbose "Error details: $documentsListItems"
            Pause
            exit 1
        }
        $documentsListItemsObj = $documentsListItems | ConvertFrom-Json
        
        # Get the eTag for the folder, if empty, use root folder
        $rawETag = $documentsListItemsObj.value | Where-Object {$_.webUrl -eq $folder.webUrl} | Select-Object -ExpandProperty eTag
        
        # Process the eTag to extract just the ID part (removing quotes and ",1" suffix)
        if ($rawETag) {
            if ($rawETag -match '"([^"]+),\d+"' -or $rawETag -match '"([^"]+)"') {
                # Extract the GUID part from the eTag, removing quotes and suffix
                $cleanETag = $Matches[1]
                Write-Verbose "Extracted clean eTag: $cleanETag from raw eTag: $rawETag"
            } else {
                # Fallback if pattern doesn't match
                $cleanETag = $rawETag -replace '"', '' -replace ',\d+$', ''
                Write-Verbose "Cleaned eTag using replace: $cleanETag from raw eTag: $rawETag"
            }
        } else {
            $cleanETag = "root"
            Write-Verbose "No eTag found, using: $cleanETag"
        }
        
        Write-Verbose "Documents list ID: $documentsListId"
        Write-Verbose "Documents list raw eTag: $rawETag"
        Write-Verbose "Documents list clean eTag: $cleanETag"

        # Add the eTag to the folder object
        $folder | Add-Member -NotePropertyName "eTag" -NotePropertyValue $cleanETag -Force
        $folder | Add-Member -NotePropertyName "eTagList" -NotePropertyValue $folder.webId -Force
        $folder.eTag = $cleanETag
        $folder.eTagList = $documentsListId
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
        Write-Verbose "Failed to get eTag for folder: $($folder.FolderName)"
        Pause
        exit 1
    }
}

Write-Host "You can now edit the folder names for shortcut creation."

try {
    # Show the folder name edit form
    $editedFolders = Show-FolderNameEditForm -SelectedFolders $selectedFolders -Verbose:$SetVerbose
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to show folder name edit form."
        pause
        exit 1
    }
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

# Check if user canceled the edit form
if (-not $editedFolders -or $editedFolders.Count -eq 0) {
    Write-Error "Folder editing was canceled. Exiting..."
    pause
    exit
}

Write-Host "Associating folders with users..."

# Associate selected folders with users
foreach ($user in $users.list) {
    if (-not $user.hasAccess) {
        Write-Warning "User $($user.displayName) does not have access to any sites. Skipping folder association."
        continue
    } else {
        $user | Add-Member -NotePropertyName "folders" -NotePropertyValue @() -Force

        # Print user information
        $userInfo = $user.sites | Select-Object siteName | Format-Table -AutoSize | Out-String
        Write-Verbose "User: $($user.DisplayName) - Sites:`n$userInfo"
        
        # For each site the user has access to
        foreach ($siteId in $user.sites.Keys) {
            # Get the driveId for this site from the user's sites collection
            $siteDriveId = $user.sites[$siteId].driveId

            Write-Verbose "Processing site: $($user.sites[$siteId].siteName)"
            Write-Verbose "Site DriveId: $siteDriveId"
            
            # Find matching folders for this site's drive ID
            $matchedFolders = @()
            foreach ($folder in $editedFolders) {
                $folderDriveId = $folder.DriveId.ToString().Trim()
                $userSiteDriveId = $siteDriveId.ToString().Trim()
                
                Write-Verbose "Comparing folder '$($folder.FolderName)': '$folderDriveId' with site drive: '$userSiteDriveId'"
                
                # If the drive IDs match, add this folder to the matched set
                if ($folderDriveId -ieq $userSiteDriveId) {
                    Write-Verbose "MATCH FOUND - Adding folder: $($folder.FolderName)"
                    $matchedFolders += $folder
                }
            }
            
            # Now add all matched folders to the user's folders collection
            if ($matchedFolders.Count -gt 0) {
                $user.folders += $matchedFolders
                Write-Verbose "Added $($matchedFolders.Count) folders from site $($user.sites[$siteId].siteName) to user $($user.DisplayName)"
                
                foreach ($folder in $matchedFolders) {
                    Write-Verbose "  - Added folder: $($folder.FolderName)"
                }
            } else {
                Write-Verbose "No matching folders found for user $($user.DisplayName) in site $($user.sites[$siteId].siteName)"
            }
        }
        
        # Show summary for this user
        Write-Verbose "Total folders for $($user.DisplayName): $(($user.folders | Measure-Object).Count)"

        # Show all folders for each user
        foreach ($folder in $user.folders) {
            Write-Verbose "  - Folder: $($folder.FolderName) (DriveId: $($folder.DriveId), Path: $($folder.Path))"
        }
    }
}

Write-Host "Processing completed. Please review the selected folders and users."

# Display the edited folders in a table format
Write-Host "Edited Folders for Shortcut Creation:" -ForegroundColor Cyan
foreach ($folder in $editedFolders) {
    Write-Host ""
    Write-Host "Foldername: $($folder.FolderName)" -ForegroundColor Green
    Write-Host "Displayname: $($folder.DisplayName)" -ForegroundColor Green
    Write-Host "Sharepoint-URL: $($folder.siteWebUrl)" -ForegroundColor Green
    Write-Host ""
}

# Show count summary
Write-Host "$($editedFolders.Count) folders will be used for shortcut creation" -ForegroundColor Green
Write-Host ""

# Show count and each folder name for each user
foreach ($user in $users.list) {
    Write-Host "User: $($user.DisplayName) ($($user.Mail))" -ForegroundColor Cyan
    Write-Host "Total folders: $($user.folders.Count)" -ForegroundColor Green
    Write-Host "Folders:" -ForegroundColor Cyan
    
    foreach ($folder in $user.folders) {
        Write-Host "  - $($folder.DisplayName)" -ForegroundColor Green
    }
    
    Write-Host ""
}

Write-Host "Obtaining WebView2 cookies for Admin Center login..."

# Grab Cookie from Admin Center Login
try {
    $acCookie = Invoke-Expression -Command "pwsh.exe -ExecutionPolicy Bypass -File `"$workpath\Get-WebView2Cookies.ps1`""
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Error: Unable to get WebView2 cookies."
        pause
        exit 1
    }
    if (-not $acCookie) {
        Write-Error "Error: No cookies found. Please ensure you are logged into the Admin Center."
        pause
        exit 1
    }
    if ($acCookie -like "*Error:*") {
        Write-Error "Error: $acCookie"
        pause
        exit 1
    }
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    pause
    exit 1
}

Write-Verbose "WebView2 cookies obtained successfully."

Write-Host "Processing OneDrive access for each user..."

# Request Onedrive-Access for each user
foreach ($user in $users.list) {
    if (-not $user.hasAccess) {
        Write-Warning "User $($user.displayName) does not have access to any sites. Skipping OneDrive access check and adding folders."
        continue
    } else {
        Write-Verbose "Processing user: $($user.DisplayName) ($($user.Mail))"
        
        $user | Add-Member -NotePropertyName "DriveAccess" -NotePropertyValue $null -Force

        # Try to get permissions for each user's OneDrive
        try {
            $userDrive = Send-GraphRequest -Method GET -Uri "https://admin.microsoft.com/admin/api/users/accessodb?upn=$($user.Mail)" -Cookie "$acCookie" -Verbose:$SetVerbose
            if ($LASTEXITCODE -ne 0) {
                Write-Error "Error: Unable to get OneDrive access for user $($user.DisplayName)."
                Write-Verbose "Error details: $userDrive"
                $user.DriveAccess = $false
                continue
            }

            # Remove "" from the response
            $userDrive = $userDrive -replace '"', ''

            Write-Verbose "User's $($user.DisplayName) ($($user.Mail)) OneDrive access: $userDrive"

            # Test access to the user's OneDrive
            # Get OneDrive's drive ID
            if ($null -eq $user.DriveAccess) {
                $userDrive = Send-GraphRequest -Method GET -Uri "$(([uri]::new("$graphEndpoint/users/$($user.id)/drives?filter=name eq 'OneDrive'")).AbsoluteUri)" -AccessToken "$accessToken" -Verbose:$SetVerbose
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Error: Unable to get OneDrive for user $($user.DisplayName)."
                    Write-Verbose "Error details: $userDrive"
                    $user.DriveAccess = $false
                    continue
                } else {
                    $userDriveObj = $userDrive | ConvertFrom-Json
                    $userDriveId = $userDriveObj.value | Where-Object {$_.name -eq "OneDrive"} | Select-Object -ExpandProperty id
                }
            }
            if (-not $userDriveId) {
                Write-Warning "No 'OneDrive' found for user: $($user.displayName) ($($user.Mail))"
                $user.DriveAccess = $false
                continue
            } else {
                Write-Verbose "User $($user.DisplayName) ($($user.Mail)) has OneDrive with ID: $userDriveId"
                $user | Add-Member -NotePropertyName "OneDriveId" -NotePropertyValue $userDriveId -Force
                $user.OneDriveId = $userDriveId
            }

            if ($null -eq $user.DriveAccess) {
                $userDriveRoot = Send-GraphRequest -Method GET -Uri "$graphEndpoint/drives/$($user.OneDriveId)/root" -AccessToken "$accessToken" -Verbose:$SetVerbose
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Error: Unable to get OneDrive root for user $($user.DisplayName)."
                    Write-Verbose "Error details: $userDriveRoot"
                    $user.DriveAccess = $false
                    continue
                } else {
                    $userDriveRootObj = $userDriveRoot | ConvertFrom-Json
                    Write-Verbose "User $($user.DisplayName) ($($user.Mail)) OneDrive root: $($userDriveRootObj.webUrl)"
                    $user.DriveAccess = $true
                }
            }
        } catch {
            Write-Error "Error: $($_.Exception.Message)"
            Write-Error "Failed to get OneDrive access for user $($user.DisplayName)."
            continue
        }

        if ($user.DriveAccess) {
            Write-Host "Attempt to get access to the user's OneDrive for user $($user.DisplayName) ($($user.Mail)) was successful." -ForegroundColor Green
            if ($user.folders.Count -gt 0) {
                Write-Host "Adding folders to user $($user.DisplayName)'s ($($user.Mail)) OneDrive..." -ForegroundColor Green
                foreach ($folder in $user.folders) {
                    Write-Host "  - Adding folder: $($folder.DisplayName) to OneDrive" -ForegroundColor Green
                    
                    # Create folder object using PowerShell hashtable instead of a raw JSON string
                    $folderObject = @{
                        name = $folder.FolderName
                        remoteItem = @{
                            sharepointIds = @{
                                listId = $folder.eTagList
                                listItemUniqueId = $folder.eTag
                                siteId = $folder.DomainSiteId
                                siteUrl = $folder.SiteWebUrl
                                webId = $folder.WebId
                            }
                        }
                        "@microsoft.graph.conflictBehavior" = "rename"
                    }
                    
                    # Convert to JSON
                    $POSTBody = $folderObject | ConvertTo-Json -Depth 10

                    try {
                        # Send the request to create the folder
                        $response = Send-GraphRequest -Method POST -Uri "$graphEndpoint/drives/$($user.OneDriveId)/root/children" -AccessToken "$accessToken" -Body $POSTBody -Verbose:$SetVerbose
                        if ($LASTEXITCODE -ne 0) {
                            Write-Error "Error: Unable to create folder $($folder.FolderName) in OneDrive for user $($user.DisplayName)."
                        }
                        $responseObj = $response | ConvertFrom-Json
                        
                        # Check if the response contains an error object
                        if ($responseObj.error) {
                            # Check for specific error codes
                            if ($responseObj.error.innerError.code -eq "shortcutAlreadyExists") {
                                Write-Warning "Shortcut already exists: $($folder.FolderName) - Creation not possible - Skipping"
                            }
                            elseif ($responseObj.error.innerError.code -eq "nestedAscendantShortcutExists") {
                                Write-Warning "A parent shortcut already exists that includes $($folder.FolderName) - Creation not possible - Skipping"
                                Write-Warning "Manaully remove the parent shortcut to create this folder."
                            }
                            elseif ($responseObj.error.innerError.code -eq "nestedDescendantShortcutExists") {
                                Write-Warning "A child shortcut already exists that includes $($folder.FolderName) - Creation not possible - Skipping"
                                Write-Warning "Manaully remove the child shortcut to create this folder."
                            }
                            else {
                                Write-Error "Error: $($responseObj.error.message)"
                                Write-Error "Failed to create folder $($folder.FolderName) in OneDrive for user $($user.DisplayName)."
                            }
                        } else {
                            # Create rename object using PowerShell hashtable
                            $renameObject = @{
                                name = $folder.DisplayName
                            }
                            
                            # Convert to JSON
                            $POSTBody = $renameObject | ConvertTo-Json

                            # Send the request to rename the folder
                            $response = Send-GraphRequest -Method PATCH -Uri "$graphEndpoint/drives/$($user.OneDriveId)/items/$($responseObj.id)" -AccessToken "$accessToken" -Body $POSTBody -Verbose:$SetVerbose
                            if ($LASTEXITCODE -ne 0) {
                                Write-Error "Error: Unable to rename folder $($folder.FolderName) in OneDrive for user $($user.DisplayName)."
                            }
                            $responseObj = $response | ConvertFrom-Json
                            if ($responseObj.error) {
                                if ($responseObj.error.code -eq "nameAlreadyExists") {
                                    Write-Error "Folder name already exists: $($folder.DisplayName) - Renaming not possible"
                                }
                                else {
                                    Write-Error "Error: $($responseObj.error.message)"
                                    Write-Error "Failed to rename folder $($folder.FolderName) in OneDrive for user $($user.DisplayName)."
                                }
                            }
                        }
                    } catch {
                        Write-Error "Error: $($_.Exception.Message)"
                        Write-Error "Failed to create folder $($folder.FolderName) in OneDrive for user $($user.DisplayName)."
                        continue
                    }
                }
            } else {
                Write-Host "No folders found for user $($user.DisplayName) ($($user.Mail)) in OneDrive." -ForegroundColor Yellow
            }
        } else {
            Write-Warning "Skipping User $($user.DisplayName) ($($user.Mail)) - Either no Onedrive access or no OneDrive found." -ForegroundColor Red
        }
    }
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