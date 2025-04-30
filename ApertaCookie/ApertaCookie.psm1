using namespace "System.Security.Cryptography"

<#
.SYNOPSIS
    Copies specified cookies sqlite database to temp directory
.DESCRIPTION
    Copies specified cookies sqlite database to temp directory to avoid issues related to a browser having a file lock on the original database.
.EXAMPLE
    Copy-CookieDBToTemp -SQLitePath $pathToCookieDB

    Copies browser cookie SQLite DB to temp dir location.
.PARAMETER SQLitePath
    Path of SQLite Cookie DB to copy
.OUTPUTS
    System.String
    -or-
    System.Boolean
.NOTES
    Author: Jake Morrison - @jakemorrison - https://www.techthoughts.info/
    This is a necessary action to prevent issues where the browser has a lock on the orignal databse.
    In testing I noticed that Linux was especially sensitive to this.
    Although I also saw the freeze behavior on Windows when querying FireFox.
    This copy action shifts ApertaCookie to querying the copy instead of the primary database.
.COMPONENT
    ApertaCookie
#>
function Copy-CookieDBToTemp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Path of SQLite Cookie DB to copy')]
        [string]
        $SQLitePath
    )

    $tempDir = [System.IO.Path]::GetTempPath()
    $tempPath = $tempDir + 'apertacookie'
    Write-Verbose -Message ('Evaluating temp location: {0}' -f $tempPath)

    if (-not(Test-Path -Path $tempPath)) {
        Write-Verbose -Message 'Creating ApertaCookie directory in temp location...'
        try {
            $newItemSplat = @{
                ItemType    = 'Directory'
                Force       = $true
                Path        = $tempPath
                ErrorAction = 'Stop'
            }
            New-Item @newItemSplat | Out-Null
            Write-Verbose -Message 'Temp directory created.'
        }
        catch {
            Write-Error $_
            return $false
        }

    }
    else {
        Write-Verbose -Message 'Temp path confirmed.'
    }

    Write-Verbose -Message 'Initiating Cookie DB copy...'
    Write-Verbose -Message ('Copying from: {0}' -f $SQLitePath)
    Write-Verbose -Message ('Copying to: {0}' -f $tempPath)

    try {
        $copySplat = @{
            Path        = $SQLitePath
            Destination = $tempPath
            Force       = $true
            Confirm     = $false
            PassThru    = $true
            ErrorAction = 'Stop'
        }
        $results = Copy-Item @copySplat
        $fullName = $results.FullName
    }
    catch {
        Write-Error $_
        return $false
    }

    Write-Verbose -Message ('Returning full copy path: {0}' -f $fullName)

    return $fullName

} #Copy-CookieDBToTemp



<#
.SYNOPSIS
    Retrieves the FireFox profile that has been written to most recently
.DESCRIPTION
    FireFox can have multiple profiles. This function finds the profile that has been written to most recently. This is likely a good indicator of the active (primary) profile.
.EXAMPLE
    Get-FireFoxProfilePath -Path $pathToFireFoxProfiles

    Returns the FireFox profile path that has been written to most recently.
.PARAMETER Path
    Path to FireFox Profiles
.OUTPUTS
    System.String
    -or
    System.Boolean
.NOTES
    Author: Jake Morrison - @jakemorrison - https://www.techthoughts.info/
    This isn't perfect. Open to better suggestions.
.COMPONENT
    ApertaCookie
#>
function Get-FireFoxProfilePath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Path to FireFox Profiles')]
        [string]
        $Path
    )

    Write-Verbose -Message ('Evaluating {0}' -f $Path)

    try {
        $itemSplat = @{
            Path        = $Path
            ErrorAction = 'Stop'
        }
        $recentProfile = Get-ChildItem @itemSplat | Where-Object { $_.PSIsContainer -and $_.Name -match '\.default' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        Write-Verbose $recentProfile
    }
    catch {
        Write-Error $_
        return $false
    }

    Write-Verbose -Message ('Most recent FireFox profile: {0}' -f $recentProfile.FullName)

    return $recentProfile.FullName

} #Get-FireFoxProfilePath


<#
.SYNOPSIS
    Returns the correct path to the SQLite Cookie database and SQLite Table Name based on OS and Browser
.DESCRIPTION
    Evaluates the provided browser and OS and returns the known location for the SQLite Cookie Database and SQLite Table Name
.EXAMPLE
    Get-OSCookieInfo -Browser 'Chrome'

    Returns SQLite Cookie Database path for OS that function is run on as well as SQLite Table Name
.PARAMETER Browser
    Browser choice
.OUTPUTS
    System.Management.Automation.PSCustomObject
.NOTES
    Author: Jake Morrison - @jakemorrison - https://www.techthoughts.info/
.COMPONENT
    ApertaCookie
#>
function Get-OSCookieInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Browser choice')]
        [ValidateSet('Edge', 'Chrome', 'FireFox', 'CustomChrome', 'CustomFireFox')]
        [string]
        $Browser,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Custom path to Chrome or FireFox Local State')]
        [string]
        $customPath
    )

    Write-Verbose -Message ('Browser: {0}' -f $Browser)

    if ($IsWindows) {
        Write-Verbose -Message 'Windows Detected'
        switch ($Browser) {
            Edge {
                $sqlPath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Network\Cookies"
                $tableName = 'cookies'
            }
            Chrome {
                $sqlPath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Network\Cookies"
                $tableName = 'cookies'
            }
            CustomChrome {
                $sqlPath = "$customPath\Default\Network\Cookies"
                $tableName = 'cookies'
            }
            FireFox {
                # TODO: check database lock. you may have to copy the file
                $profilePath = Get-FireFoxProfilePath -Path "$env:APPDATA\Mozilla\Firefox\Profiles"
                if ($profilePath) {
                    $sqlPath = "$profilePath\cookies.sqlite"
                    $tableName = 'moz_cookies'
                }
                else {
                    # TODO: ERROR CONTROL
                }
            }
            CustomFireFox {
                # TODO: check database lock. you may have to copy the file
                $profilePath = Get-FireFoxProfilePath -Path "$customPath\Profiles"
                if ($profilePath) {
                    $sqlPath = "$profilePath\cookies.sqlite"
                    $tableName = 'moz_cookies'
                }
                else {
                    # TODO: ERROR CONTROL
                }
            }
        }
    } #if_Windows
    elseif ($IsLinux) {
        Write-Verbose -Message 'Linux Detected'
        switch ($Browser) {
            Edge {
                $sqlPath = "$env:HOME/.config/microsoft-edge-beta/Default/Network/Cookies"
                $tableName = 'cookies'
            }
            Chrome {
                $sqlPath = "$env:HOME/.config/google-chrome/Default/Network/Cookies"
                $tableName = 'cookies'
            }
            CustomChrome {
                $sqlPath = "$customPath/Default/Network/Cookies"
                $tableName = 'cookies'
            }
            FireFox {
                $profilePath = Get-FireFoxProfilePath -Path "$env:HOME/.mozilla/firefox"
                if ($profilePath) {
                    $sqlPath = "$profilePath/cookies.sqlite"
                    $tableName = 'moz_cookies'
                }
                else {
                    # TODO: ERROR CONTROL
                }
            }
            CustomFireFox {
                $profilePath = Get-FireFoxProfilePath -Path "$customPath"
                if ($profilePath) {
                    $sqlPath = "$profilePath/cookies.sqlite"
                    $tableName = 'moz_cookies'
                }
                else {
                    # TODO: ERROR CONTROL
                }
            }
        } #elseif_Linux
    }
    elseif ($IsMacOS) {
        Write-Verbose -Message 'OSX Detected'
        switch ($Browser) {
            Edge {
                $sqlPath = "$env:HOME/Library/Application Support/Microsoft Edge/Default/Network/Cookies"
                $tableName = 'cookies'
            }
            Chrome {
                $sqlPath = "$env:HOME/Library/Application Support/Google/Chrome/Default/Network/Cookies"
                $tableName = 'cookies'
            }
            CustomChrome {
                $sqlPath = "$customPath/Default/Network/Cookies"
                $tableName = 'cookies'
            }
            FireFox {
                $profilePath = Get-FireFoxProfilePath -Path "$env:HOME/Library/Application Support/Firefox/Profiles"
                if ($profilePath) {
                    $sqlPath = "$profilePath/cookies.sqlite"
                    $tableName = 'moz_cookies'
                }
                else {
                    # TODO: ERROR CONTROL
                }
            }
            CustomFireFox {
                $profilePath = Get-FireFoxProfilePath -Path "$customPath"
                if ($profilePath) {
                    $sqlPath = "$profilePath/cookies.sqlite"
                    $tableName = 'moz_cookies'
                }
                else {
                    # TODO: ERROR CONTROL
                }
            }
        } #elseif_OSX
    }

    Write-Verbose -Message ('SQLite Path: {0}' -f $sqlPath)
    Write-Verbose -Message ('Table Name: {0}' -f $tableName)

    $obj = [PSCustomObject]@{
        SQLitePath = $sqlPath
        TableName  = $tableName
    }

    return $obj
}
<#
.SYNOPSIS
    Retrievs the primary cookie decryption key for Edge or Chrome cookies
.DESCRIPTION
    Queries the Local State and uses the currentuser context to decrypt the cookies key which can be used to decrypt cookies. This decrypt context can then be loaded into an AesGcm in the context of $Script:GCMKey
.EXAMPLE
    Get-WindowsCookieDecryptKey -Browser Edge

    Retrieves the cookie decryption key using the currentuser context for Edge
.PARAMETER Browser
    Browser choice
.OUTPUTS
    System.Byte
.NOTES
    Author: Jake Morrison - @jakemorrison - https://www.techthoughts.info/
    This was very hard to make.
    $Script:GCMKey should be used for decrypting cookies
.COMPONENT
    ApertaCookie
#>
function Get-WindowsCookieDecryptKey {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Browser choice')]
        [ValidateSet('Edge', 'Chrome', 'FireFox', 'CustomChrome', 'CustomFireFox')]
        [string]
        $Browser,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Custom path to Chrome or FireFox Local State')]
        [string]
        $customPath

    )

    Write-Verbose -Message ('Browser: {0}' -f $Browser)

    switch ($Browser) {
        Edge {
            $statePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Local State"
        }
        Chrome {
            $statePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Local State"
        }
        CustomChrome {
            $statePath = "$customPath\Local State"
        }
        FireFox {
            Write-Verbose -Message 'No state decryption needed for FireFox. Skipping'
            return
        }
        CustomFireFox {
            Write-Verbose -Message 'No state decryption needed for FireFox. Skipping'
            return
        }
    }

    Write-Verbose -Message ('Retrieving content from {0}' -f $statePath)
    try {
        $contentSplat = @{
            Path        = $statePath
            ErrorAction = 'Stop'
        }
        $cookiesKeyEncBaseSixtyFourRaw = Get-Content @contentSplat
    }
    catch {
        Write-Error $_
        return $false
    }

    Write-Verbose -Message 'Converting from JSON...'
    $cookiesKeyEncBaseSixtyFour = ($cookiesKeyEncBaseSixtyFourRaw | ConvertFrom-Json -AsHashtable).'os_crypt'.'encrypted_key'
    Write-Verbose -Message 'Converting from Base64...'
    $cookiesKeyEnc = [System.Convert]::FromBase64String($cookiesKeyEncBaseSixtyFour) | Select-Object -Skip 5  # Magic number 5

    Write-Verbose -Message 'Running Unprotect...'
    try {
        $cookiesKey = [System.Security.Cryptography.ProtectedData]::Unprotect($cookiesKeyEnc, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    }
    catch {
        Write-Error $_
        return $false
    }

    return $cookiesKey

} #Get-WindowsCookieDecryptKey


<#
.SYNOPSIS
    Creates System.Security.Cryptography.AesGcm with provided key byte
.DESCRIPTION
    Creates System.Security.Cryptography.AesGcm with provided key byte from the decrypted Local State
.EXAMPLE
    New-WindowsAesGcm -WinDecrypt $key
.PARAMETER WinDecrypt
    Windows Local State Key
.OUTPUTS
    System.Bool
.NOTES
    Author: Jake Morrison - @jakemorrison - https://www.techthoughts.info/
    This was very hard to make.
    $Script:GCMKey should be used for decrypting cookies
    Some of this code was inspired from Read-Chromium by James O'Neill:
        https://www.powershellgallery.com/packages/Read-Chromium/1.0.0/Content/Read-Chromium.ps1
.COMPONENT
    ApertaCookie
#>
function New-WindowsAesGcm {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Windows Local State Key')]
        [System.Object]
        $WinDecrypt

    )

    Write-Verbose -Message 'Creating AesGcm...'

    try {
        $Script:GCMKey = [AesGcm]::new($WinDecrypt)
        Write-Verbose -Message 'Decrypt Key loaded into script variable!'
    }
    catch {
        Write-Error $_
        return $false
    }

    return $true
} #New-WindowsAesGcm


<#
.SYNOPSIS
    Takes in cookie encrypted byte encrypted_value and returns decrypted cookie value
.DESCRIPTION
    Using the GCMKey decrypt key returns decrypted cookie value from a provided cookie encrypted_value byte value
.EXAMPLE
    Unprotect-Cookie -Cookie $cookies[0].encrypted_value -Verbose

    Decrypts the provided cookie byte
.PARAMETER Cookie
    Encrypted cookie value in byte format
.OUTPUTS
    System.String
.NOTES
    Author: Jake Morrison - @jakemorrison - https://www.techthoughts.info/
    Use GCM decrytpion if ciphertext starts "V10" & GCMKey exists, else try ProtectedData.unprotect
    Requires Get-CookieDecryptKey to have already been run to load $Script:GCMKey
    Some of this code was inspired from Read-Chromium by James O'Neill:
        https://www.powershellgallery.com/packages/Read-Chromium/1.0.0/Content/Read-Chromium.ps1
.COMPONENT
    ApertaCookie
#>
function Unprotect-Cookie {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Encrypted cookie value in byte format')]
        [System.Object]
        $Cookie,
        [Parameter(Mandatory = $false)]
        [int]
        $DBVersion
    )

    Write-Verbose -Message 'Decrypting cookie...'

    try {
        if ($Script:GCMKey -and [string]::new($Cookie[0..2]) -match "v1\d") {
            Write-Verbose -Message 'AesGcm decrypt'
            #Ciphertext bytes run 0-2="V10"; 3-14=12_byte_IV; 15 to len-17=payload; final-16=16_byte_auth_tag
            $output = [System.Byte[]]::new($Cookie.length - 31) # same length as payload.

            #_____________________________________________________________________________________________
            # https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.aesgcm?view=net-5.0
            $Script:GCMKey.Decrypt(
                $Cookie[3..14],
                $Cookie[15..($Cookie.Length - 17)],
                $Cookie[-16..-1],
                $output,
                $null)

            #_____________________________________________________________________________________________

            # Remove SHA256(host_key) prefix for DBVersion >= 24
            if ($DBVersion -ge 24) {
                $payload = $output[32..($output.Length-1)]
            } else {
                $payload = $output
            }

            # Remove PKCS#7 padding
            $padLen = $payload[-1]
            if ($padLen -ge 1 -and $padLen -le 16) {
                $payload = $payload[0..($payload.Length - $padLen - 1)]
            }

            return [System.Text.Encoding]::UTF8.GetString($payload)

        }
        else {
            Write-Verbose -Message 'Attempting CurrentUser decryption method'
            [string]::new([ProtectedData]::Unprotect($Cookie, $null, 'CurrentUser'))
        }
    }
    catch {
        Write-Warning $_
    }
} #Unprotect-Cookie


<#
.EXTERNALHELP ApertaCookie-help.xml
#>
function Convert-CookieTime {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Chrome Based Time')]
        [Int64]
        $CookieTime,
        [Parameter(HelpMessage = 'Specify switch if converting from FireFox time')]
        [switch]
        $FireFoxTime
    )

    # 01/01/1601 00:00:00 - Chrome time date to vet against
    # 13267233550477440 - time input to expect

    # 01/01/1970 00:00:00 - firefox time date to vet against
    # 1616989552356002 - firefox time input

    Write-Verbose -Message 'Converting cookie time to dateTime!'

    if ($FireFoxTime) {
        Write-Verbose -Message 'FireFox detected. Taking us back to 1970! Groovy!'
        [datetime]$baseTime = '01/01/1970 00:00:00'
    }
    else {
        Write-Verbose -Message 'Taking us back to 1601!'
        [datetime]$baseTime = '01/01/1601 00:00:00'
    }

    $seconds = $CookieTime / 1000000
    [dateTime]$convertedTime = $baseTime.AddSeconds($seconds)

    return $convertedTime
} #Convert-CookieTime



<#
.EXTERNALHELP ApertaCookie-help.xml
#>
function Get-DecryptedCookiesInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Browser choice')]
        [ValidateSet('Edge', 'Chrome', 'FireFox', 'CustomChrome', 'CustomFireFox')]
        [string]
        $Browser,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Domain to search for in Cookie Database')]
        [string]
        $DomainName,
        [Parameter(HelpMessage = 'If specified cookies are loaded into a Microsoft.PowerShell.Commands.WebRequestSession and returned for session use.')]
        [switch]
        $WebSession,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Custom path to Chrome or FireFox Local State')]
        [string]
        $customPath
    )

    Write-Verbose -Message 'Retrieving raw cookies from SQLite database...'
    try {
        if ($Browser -eq 'CustomChrome' -or $Browser -eq 'CustomFireFox') {
            if ($DomainName) {
                $cookies = Get-RawCookiesFromDB -Browser $Browser -DomainName $DomainName -customPath $customPath -ErrorAction 'Stop'
            }
            else {
                $cookies = Get-RawCookiesFromDB -Browser $Browser -customPath $customPath -ErrorAction 'Stop'
            }
        } else {
            if ($DomainName) {
                $cookies = Get-RawCookiesFromDB -Browser $Browser -DomainName $DomainName -ErrorAction 'Stop'
            }
            else {
                $cookies = Get-RawCookiesFromDB -Browser $Browser -ErrorAction 'Stop'
            }
        }
    }
    catch {
        throw $_
    }

    if ($null -eq $cookies) {
        Write-Warning -Message 'No cookies were returned from the SQLite database!'
        return
    }

    if ($Browser -eq 'FireFox') {
        Write-Verbose -Message 'FireFox specified. No cookie decryption is necessary.'
        if ($WebSession) {
            Write-Verbose -Message ('Cookie Count: {0}' -f $cookies.Count)
            if ($cookies.Count -gt 300) {
                throw 'Only up to 300 cookies can be loaded into a WebSession.'
            }

            Write-Verbose -Message 'WebSession specified. Loading cookies into WebSession.'
            $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

            foreach ($cookie in $cookies) {
                $newCookie = New-Object System.Net.Cookie

                $newCookie.Name = $cookie.name
                $newCookie.Value = $cookie.value
                $newCookie.Domain = $cookie.host

                try {
                    $session.Cookies.Add($newCookie)
                }
                catch {
                    Write-Warning -Message "$($cookie.name) could not be loaded into the session. Skipping."
                }

            } #foreach_cookies

            return $session

        } #if_websession
        else {

            return $cookies

        } #else_websession
    } #if_FireFox
    
    if ($Browser -eq 'CustomFireFox') {
        Write-Verbose -Message 'Custom FireFox specified. No cookie decryption is necessary.'
        if ($WebSession) {
            Write-Verbose -Message ('Cookie Count: {0}' -f $cookies.Count)
            if ($cookies.Count -gt 300) {
                throw 'Only up to 300 cookies can be loaded into a WebSession.'
            }

            Write-Verbose -Message 'WebSession specified. Loading cookies into WebSession.'
            $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

            foreach ($cookie in $cookies) {
                $newCookie = New-Object System.Net.Cookie

                $newCookie.Name = $cookie.name
                $newCookie.Value = $cookie.value
                $newCookie.Domain = $cookie.host

                try {
                    $session.Cookies.Add($newCookie)
                }
                catch {
                    Write-Warning -Message "$($cookie.name) could not be loaded into the session. Skipping."
                }

            } #foreach_cookies

            return $session

        } #if_websession
        else {

            return $cookies

        } #else_websession
    } #if_CustomFireFox


    Write-Verbose -Message 'Getting Cookie decryption key...'

    if ($IsWindows) {
        if ($Browser -eq 'CustomChrome' -or $Browser -eq 'CustomFireFox') {
            Write-Verbose -Message 'Windows detected...'
            $cookiesKey = Get-WindowsCookieDecryptKey -Browser $Browser -customPath $customPath
            if ($cookiesKey -eq $false) {
                throw 'Cookie decryption key not retrieved successfully'
            }
            #sets the script level decryption key
            $decryptEval = New-WindowsAesGcm -WinDecrypt $cookiesKey
            if ($decryptEval -ne $true) {
                throw 'AesGcm Cookie decryption key not created successfully'
            }
        } else {
            Write-Verbose -Message 'Windows detected...'
            $cookiesKey = Get-WindowsCookieDecryptKey -Browser $Browser
            if ($cookiesKey -eq $false) {
                throw 'Cookie decryption key not retrieved successfully'
            }
            #sets the script level decryption key
            $decryptEval = New-WindowsAesGcm -WinDecrypt $cookiesKey
            if ($decryptEval -ne $true) {
                throw 'AesGcm Cookie decryption key not created successfully'
            }
        }
    }
    elseif ($IsLinux) {
        Write-Verbose -Message 'Linux detected...'
        #TODO: Add Linux cookie decrypt support
        Write-Warning -Message 'Linux cookie decryption is not currently supported'
        return
    }
    elseif ($IsMacOS) {
        Write-Verbose -Message 'MacOS detected...'
        #TODO: Add MacOS support
        Write-Warning -Message 'MacOS cookie decryption is not currently supported'
        return
    }
    else {
        throw 'Unsupported OS. Windows / Linux / OSX'
    }

    Write-Verbose -Message 'Decrypting cookies...'
    if ($Browser -eq 'CustomChrome') {
        Write-Verbose -Message "Custom Chrome detected. Getting DB version..."
        $DBVersion = Get-ChromeCookieDBVersion -Browser $Browser -customPath $customPath
    }
    if ($Browser -eq 'Chrome' -or $Browser -eq 'Edge') {
        Write-Verbose -Message "Chrome detected. Getting DB version..."
        $DBVersion = Get-ChromeCookieDBVersion -Browser $Browser
    }
    foreach ($cookie in $cookies) {
        if ($Browser -eq 'Chrome' -or $Browser -eq 'Edge' -or $Browser -eq 'CustomChrome') {
            $temp = $null
            $temp = Unprotect-Cookie -Cookie $cookie.encrypted_value -DBVersion $DBVersion
            $cookie | Add-Member -NotePropertyName decrypted_value -NotePropertyValue $temp
        } else {
            $temp = $null
            $temp = Unprotect-Cookie -Cookie $cookie.encrypted_value 
            $cookie | Add-Member -NotePropertyName decrypted_value -NotePropertyValue $temp
        }
    }
    Write-Verbose -Message 'Cookies decryption completed.'

    if ($WebSession) {
        Write-Verbose -Message ('Cookie Count: {0}' -f $cookies.Count)
        if ($cookies.Count -gt 300) {
            throw 'Only up to 300 cookies can be loaded into a WebSession.'
        }

        Write-Verbose -Message 'WebSession specified. Loading cookies into WebSession.'
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

        foreach ($cookie in $cookies) {
            $newCookie = New-Object System.Net.Cookie

            $newCookie.Name = $cookie.name
            $newCookie.Value = $cookie.decrypted_value
            $newCookie.Domain = $cookie.host_key

            try {
                $session.Cookies.Add($newCookie)
            }
            catch {
                Write-Warning -Message "$($cookie.name) could not be loaded into the session. Skipping."
            }

        } #foreach_cookies

        return $session

    } #if_websession
    else {
        return $cookies
    } #else_websession

} #Get-DecryptedCookiesInfo



<#
.EXTERNALHELP ApertaCookie-help.xml
#>
function Get-RawCookiesFromDB {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Browser choice')]
        [ValidateSet('Edge', 'Chrome', 'FireFox', 'CustomChrome', 'CustomFireFox')]
        [string]
        $Browser,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Domain to search for in Cookie Database')]
        [string]
        $DomainName,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Custom path to Chrome or FireFox Local State')]
        [string]
        $customPath
    )

    $cookies = $null

    if ($Browser -eq 'CustomChrome' -or $Browser -eq 'CustomFireFox') {
        $osCookieInfo = Get-OSCookieInfo -Browser $Browser -customPath $customPath
        if ($null -eq $osCookieInfo) {
            Write-Warning 'OS Cookie information was not located'
            return
        }
    } else {
        $osCookieInfo = Get-OSCookieInfo -Browser $Browser
        if ($null -eq $osCookieInfo) {
            Write-Warning 'OS Cookie information was not located'
            return
        }
    }

    $sqlPath = $osCookieInfo.SQLitePath
    $tableName = $osCookieInfo.TableName

    $copySQLPath = Copy-CookieDBToTemp -SQLitePath $sqlPath

    Write-Verbose -Message 'Copying sqlite db to temp location for query...'
    if ($copySQLPath -eq $false) {
        Write-Warning 'OS Cookie database could not be copied for query!'
        return
    }
    else {
        Write-Verbose -Message 'Copy completed.'
    }

    if ($DomainName) {
        switch ($Browser) {
            FireFox {
                $query = "SELECT `"_rowid_`",* FROM `"main`".`"$tableName`" WHERE `"host`" LIKE '%$DomainName%' ESCAPE '\' LIMIT 0, 49999;"
            }
            CustomFireFox {
                $query = "SELECT `"_rowid_`",* FROM `"main`".`"$tableName`" WHERE `"host`" LIKE '%$DomainName%' ESCAPE '\' LIMIT 0, 49999;"
            }
            Default {
                $query = "SELECT `"_rowid_`",* FROM `"main`".`"$tableName`" WHERE `"host_key`" LIKE '%$DomainName%' ESCAPE '\' LIMIT 0, 49999;"
            }
        }

    }
    else {
        $query = "SELECT `"_rowid_`",* FROM `"main`".`"$tableName`" '\' LIMIT 0, 49999;"
    }

    Write-Verbose -Message ('Establishing SQLite connection to: {0}' -f $copySQLPath)
    try {
        $cookiesSQL = New-SQLiteConnection -DataSource $copySQLPath -ErrorAction 'Stop'
    }
    catch {
        if ($copySQLPath -ne $false) {
            Write-Verbose -Message 'Attempting sqlite db copy cleanup...'
            Remove-Item -Path $copySQLPath -Confirm:$false -Force -ErrorAction 'SilentlyContinue'
        }
        throw $_
    }

    # examples of things that can be done with the connection:
    # $cookiesSQL.GetSchema()
    # $cookiesSQL.GetSchema("Tables")

    Write-Verbose -Message ('Running query {0} against {1}' -f $query, $sqlPath)
    try {
        $dataSource = $cookiesSQL.FileName
        $sqlSplat = @{
            Query       = $query
            DataSource  = $dataSource
            ErrorAction = 'Stop'
        }
        $cookies = Invoke-SqliteQuery @sqlSplat
    }
    catch {
        if ($copySQLPath -ne $false) {
            Write-Verbose -Message 'Attempting sqlite db copy cleanup...'
            Remove-Item -Path $copySQLPath -Confirm:$false -Force -ErrorAction 'SilentlyContinue'
        }
        throw $_
    }
    finally {
        $cookiesSQL.Close()
        $cookiesSQL.Dispose()
        if ($copySQLPath -ne $false) {
            Write-Verbose -Message 'Attempting sqlite db copy cleanup...'
            Remove-Item -Path $copySQLPath -Confirm:$false -Force -ErrorAction 'SilentlyContinue'
        }
    }

    Write-Verbose -Message 'Cookies retrieved from SQLite'

    return $cookies

} #Get-RawCookiesFromDB

function Get-ChromeCookieDBVersion {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'Browser choice')]
        [ValidateSet('Edge', 'Chrome', 'CustomChrome')]
        [string]
        $Browser,
        [Parameter(Mandatory = $false,
            HelpMessage = 'Custom path to Chrome Local State')]
        [string]
        $customPath
    )

    if ($Browser -eq 'CustomChrome') {
        $osCookieInfo = Get-OSCookieInfo -Browser $Browser -customPath $customPath
        if ($null -eq $osCookieInfo) {
            Write-Warning 'OS Cookie information was not located'
            return
        }
    } else {
        $osCookieInfo = Get-OSCookieInfo -Browser $Browser
        if ($null -eq $osCookieInfo) {
            Write-Warning 'OS Cookie information was not located'
            return
        }
    }

    $sqlPath = $osCookieInfo.SQLitePath

    $copySQLPath = Copy-CookieDBToTemp -SQLitePath $sqlPath

    Write-Verbose -Message 'Copying sqlite db to temp location for query...'
    if ($copySQLPath -eq $false) {
        Write-Warning 'OS Cookie database could not be copied for query!'
        return
    }
    else {
        Write-Verbose -Message 'Copy completed.'
    }

    $query = "SELECT `"value`" FROM `"main`".`"meta`" WHERE `"key`" = 'version';"

    Write-Verbose -Message ('Establishing SQLite connection to: {0}' -f $copySQLPath)
    try {
        $cookiesSQL = New-SQLiteConnection -DataSource $copySQLPath -ErrorAction 'Stop'
    }
    catch {
        if ($copySQLPath -ne $false) {
            Write-Verbose -Message 'Attempting sqlite db copy cleanup...'
            Remove-Item -Path $copySQLPath -Confirm:$false -Force -ErrorAction 'SilentlyContinue'
        }
        throw $_
    }

    Write-Verbose -Message ('Running query {0} against {1}' -f $query, $sqlPath)
    try {
        $dataSource = $cookiesSQL.FileName
        $sqlSplat = @{
            Query       = $query
            DataSource  = $dataSource
            ErrorAction = 'Stop'
        }
        $result = Invoke-SqliteQuery @sqlSplat
    }
    catch {
        if ($copySQLPath -ne $false) {
            Write-Verbose -Message 'Attempting sqlite db copy cleanup...'
            Remove-Item -Path $copySQLPath -Confirm:$false -Force -ErrorAction 'SilentlyContinue'
        }
        throw $_
    }
    finally {
        $cookiesSQL.Close()
        $cookiesSQL.Dispose()
        if ($copySQLPath -ne $false) {
            Write-Verbose -Message 'Attempting sqlite db copy cleanup...'
            Remove-Item -Path $copySQLPath -Confirm:$false -Force -ErrorAction 'SilentlyContinue'
        }
    }

    if ($result -and $result.value) {
        return $result.value
    } else {
        Write-Warning 'Could not retrieve Chrome cookie DB version.'
        return $null
    }
}