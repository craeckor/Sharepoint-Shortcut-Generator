$workpath = $PSScriptRoot
Set-Location -Path $workpath
$url = "https://admin.microsoft.com/login"
$quiturl = "https://admin.microsoft.com"
$searchdomain = "admin.microsoft.com"
$cookiePath = "$env:TEMP\PSAuthClientWebview2Cache\EBWebView"
$returncookie = $null
Invoke-Expression -Command "$workpath\Import-Assemblies.ps1"
Import-Module -Name "$workpath\PSAuthClient\PSAuthClient.psd1" -Force
Import-Module -Name "$workpath\PSSQLLite\PSSQLite.psd1" -Force
Import-Module -Name "$workpath\ApertaCookie\ApertaCookie.psd1" -Force

try {
    Clear-WebView2Cache -Confirm:$false
} catch {
    return $($_.Exception.Message)
}

try {
    Invoke-WebView2 -uri "$url" -UrlCloseConditionRegex "$quiturl" | Out-Null
} catch {
    return $($_.Exception.Message)
}
try {
    $cookies = Get-DecryptedCookiesInfo -Browser "CustomChrome" -DomainName "$searchdomain" -customPath "$cookiePath"
} catch {
    return $($_.Exception.Message)
}


Foreach ($cookie in $cookies) {
    if ($null -eq $returncookie) {
        $returncookie = "$($cookie.name)" + "=" + "$($cookie.decrypted_value)"
    } else {
        $returncookie = "$($returncookie)" + ";" + "$($cookie.name)" + "=" + "$($cookie.decrypted_value)"
    }
}

return "$($returncookie)"