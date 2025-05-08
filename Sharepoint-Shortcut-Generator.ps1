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
    Required Scope-Permissions: files.readwrite.all user.read.all allsites.fullcontrol allsites.manage myfiles.read myfiles.write user.readwrite.all
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

function Restart-AsAdmin  {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        $argList = @(
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$PSCommandPath`"",
            "-tenantId", "`"$tenantId`"",
            "-clientId", "`"$clientId`""
        )
        
        # Pass the Verbose parameter if it was provided
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
    
    # Pass the Verbose parameter if it was provided
    if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"]) {
        $argList += "-Verbose"
    }
    
    Start-Process powershell.exe -ArgumentList $argList
    exit
}

Restart-AsAdmin

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
        
        Write-Verbose "Access token successfully obtained"
        
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

function Get-GraphRequest {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Uri,
        [Parameter(Mandatory=$true)]
        [string]$AccessToken
    )
    
    $maxRetries = 5
    $retryCount = 0
    $success = $false
    
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
            
            $rawResponse = Invoke-Expression -Command "$curlpath --connect-timeout 45 --retry 5 --retry-max-time 120 --retry-connrefused -o - -S -s -w '%{http_code}' -X GET -H `"Authorization: Bearer $AccessToken`" -H `"Content-Type: application/json`" -H `"Accept: application/json`" `"$Uri`""
            
            if ($LASTEXITCODE -ne 0) {
                Write-Verbose "Curl command failed with exit code $LASTEXITCODE"
                throw "Curl command failed with exit code $LASTEXITCODE"
            }
            
            # Extract the HTTP status code (last 3 characters)
            $httpCode = $rawResponse.Substring($rawResponse.Length - 3)
            
            # Extract the actual response body (everything except the last 3 characters)
            $responseBody = $rawResponse.Substring(0, $rawResponse.Length - 3)
            
            # Check if HTTP status code is 403 - no retry for this specific status
            if ($httpCode -eq "403") {
                # Store failed request information
                $script:requests.failed[$currentRequestId] = @{
                    Uri = $Uri
                    Timestamp = Get-Date
                    Error = "HTTP request failed with status code: 403 (Forbidden)"
                    HttpCode = $httpCode
                    Response = $responseBody
                }
                
                Write-Verbose "HTTP request #$currentRequestId failed with status code: 403 (Forbidden) - Not retrying"
                throw "HTTP request failed with status code: 403 (Forbidden)"
            }
            
            # Check if HTTP status code indicates other errors that we should retry
            if (-not ($httpCode -match "^(2\d\d)$")) {
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
            }
            
            Write-Verbose "Request #$currentRequestId completed successfully with status code: $httpCode after $retryCount retries"
            
            # Return just the response body
            return $responseBody
        }
        catch {
            $retryCount++
            
            # Check if we've reached max retries or if it's a 403 error (which we don't retry)
            if ($retryCount -gt $maxRetries -or $_.Exception.Message -match "403 \(Forbidden\)") {
                # Store failed request information if it's the final attempt
                if (-not $script:requests.failed.ContainsKey($currentRequestId)) {
                    $script:requests.failed[$currentRequestId] = @{
                        Uri = $Uri
                        Timestamp = Get-Date
                        Error = $_.Exception.Message
                        HttpCode = if ($httpCode) { $httpCode } else { $null }
                        Response = if ($responseBody) { $responseBody } else { $null }
                        RetryCount = $retryCount - 1
                    }
                }
                
                Write-Error "Error in GraphAPI request after $($retryCount-1) retries: $($_.Exception.Message)"
                return $_.Exception.Message
            }
            
            Write-Verbose "Request failed. Error: $($_.Exception.Message). Retrying..."
        }
    } while (-not $success -and $retryCount -le $maxRetries)
}

try {
    $tokenInfo = Get-OAuth2Token -TenantId $tenantId -ClientId $clientId -Verbose
    $accessToken = $tokenInfo.AccessToken
    
    Write-Verbose "Token will expire at: $($tokenInfo.ExpiresAt)"
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    pause
    exit 1
}

# Get all users in the tenant
$allUsers = Get-GraphRequest -Uri "$graphEndpoint/users" -AccessToken "$accessToken"

# Parse the JSON response
$allUsersObj = $allUsers | ConvertFrom-Json

# Generate and display the table
$usersAllTable = ($allUsersObj.value | Select-Object DisplayName, mail, id | Format-Table -AutoSize | Out-String).Trim()
Write-Verbose "All users in tenant:`n$usersAllTable"

# Only call ONCE and assign result
$selectedUsers = Show-UserSelectionForm -userList $allUsersObj.value

if (-not $selectedUsers -or $selectedUsers.Count -eq 0) {
    Write-Host "No users selected. Exiting..."
    pause
    exit
}

$users.list = $selectedUsers

# Fix for Write-Verbose formatting issues
Write-Verbose "Users object: $($users | ConvertTo-Json -Depth 2 -Compress)"
Write-Verbose "Selected users count: $($users.list.Count)"
# Format as string first, then output with Write-Verbose
$usersTable = ($users.list | Format-Table -AutoSize | Out-String).Trim()
Write-Verbose "Selected users:`n$usersTable"

$allSites = Get-GraphRequest -Uri "$graphEndpoint/sites?search=*" -AccessToken "$accessToken"

$allSitesObj = $allSites | ConvertFrom-Json

$sitesAllTable = ($allSitesObj.value | Select-Object DisplayName, WebUrl, id | Format-Table -AutoSize | Out-String).Trim()
Write-Verbose "All users in tenant:`n$sitesAllTable"

$selectedSites = Show-SiteSelectionForm -SiteList $allSitesObj.value

if (-not $selectedSites -or $selectedSites.Count -eq 0) {
    Write-Host "No sites selected. Exiting..."
    pause
    exit
}

$sites.list = $selectedSites

# Format as string first, then output with Write-Verbose
$sitesTable = ($sites.list | Format-Table -Property DisplayName, WebUrl, id -AutoSize | Out-String).Trim()
Write-Verbose "Selected sites:`n$sitesTable"

foreach ($site in $sites.list) {
    $site | Add-Member -NotePropertyName "DriveId" -NotePropertyValue $null -Force

    $drive = Get-GraphRequest -Uri "$graphEndpoint/sites/$($site.id)/drives" -AccessToken "$accessToken"

    $driveObj = $drive | ConvertFrom-Json

    $site.DriveId = $driveObj.value | Where-Object {$_.name -eq "Documents"} | Select-Object -ExpandProperty id
}

# Format as string first, then output with Write-Verbose
$sitesTable = ($sites.list | Format-Table -Property DisplayName, DriveId, WebUrl, id -AutoSize | Out-String).Trim()
Write-Verbose "Selected sites:`n$sitesTable"

foreach ($site in $sites.list) {
    $site | Add-Member -NotePropertyName "Permissions" -NotePropertyValue @{} -Force

    Write-Verbose "Getting permissions for site: $($site.displayName)"
    
    # Get User Information List
    $permissionslist = Get-GraphRequest -Uri "$(([uri]::new("$graphEndpoint/sites/$($site.id)/lists?filter=displayName eq 'User Information List'")).AbsoluteUri)" -AccessToken "$accessToken"
    $permissionslistObj = $permissionslist | ConvertFrom-Json
    $permissionslistId = $permissionslistObj.value | Where-Object {$_.displayName -eq "User Information List"} | Select-Object -ExpandProperty id
    
    if (-not $permissionslistId) {
        Write-Warning "No 'User Information List' found for site: $($site.displayName)"
        continue
    }
    
    # Get list items with all fields
    $permissions = Get-GraphRequest -Uri "$graphEndpoint/sites/$($site.id)/lists/$permissionslistId/items?expand=fields" -AccessToken "$accessToken"
    $permissionsObj = $permissions | ConvertFrom-Json
    
    # Extract only the fields property from each item
    $fieldsOnly = $permissionsObj.value | Select-Object -ExpandProperty fields
    
    # Store in the site object
    $site.Permissions = $fieldsOnly
    
    # Display first 5 items with their key properties
    Write-Verbose "Fields for first 5 users in site '$($site.displayName)':"
    $fieldsTable = ($fieldsOnly | Select-Object | 
        Select-Object EMail, Title, Name, FirstName, LastName, ContentType, Created, Modified -First 5 |
        Format-Table -AutoSize | Out-String).Trim()
    Write-Verbose $fieldsTable
    
    # Show total count
    Write-Verbose "Total users in permissions list: $($fieldsOnly.Count)"
}

foreach ($user in $users.list) {
    $user | Add-Member -NotePropertyName "sites" -NotePropertyValue @{} -Force
    $userEmail = $user.mail
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
        Write-Verbose "User $($user.displayName) has access to $siteCount sites"
    } else {
        Write-Warning "User $($user.displayName) does not have access to any of the selected sites"
    }
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
        }
    }
}

# Show the folder selection form
$selectedFolders = Show-FolderSelectionForm -DriveList $driveList -AccessToken $accessToken

# Display selected folders
if (-not $selectedFolders -or $selectedFolders.Count -eq 0) {
    Write-Host "No folders selected. Exiting..." -ForegroundColor Yellow
    pause
    exit
}

# Print selected folders
Write-Host "Selected Folders:" -ForegroundColor Cyan
$foldersTable = $selectedFolders | Format-Table -Property * -AutoSize | Out-String
Write-Host $foldersTable

# Associate selected folders with users
foreach ($user in $users.list) {
    $user | Add-Member -NotePropertyName "folders" -NotePropertyValue @() -Force

    # Print user information
    $userInfo = $user.sites | Format-Table -Property * -AutoSize | Out-String
    Write-Verbose "User: $($user.DisplayName) - Sites:`n$userInfo"
    
    # For each site the user has access to
    foreach ($siteId in $user.sites.Keys) {
        # Get the driveId for this site from the user's sites collection
        $siteDriveId = $user.sites[$siteId].driveId

        Write-Verbose "Processing site: $($user.sites[$siteId].siteName)"
        Write-Verbose "Site DriveId: $siteDriveId"
        
        # Find matching folders for this site's drive ID
        $matchedFolders = @()
        foreach ($folder in $selectedFolders) {
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
        Write-Verbose "  - Folder: $($folder.FolderName) (DriveId: $($folder.DriveId))"
    }
}

# Print summary of folder assignments
Write-Host "`nFolder Assignment Summary:" -ForegroundColor Cyan
foreach ($user in $users.list) {
    $folderCount = ($user.folders | Measure-Object).Count
    Write-Host "- $($user.DisplayName): $folderCount folders" -ForegroundColor White
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