$finalcookie = $null
Foreach ($cookie in $admincookies) {
    if ($finalcookie -eq $null) {
        $finalcookie = "$($cookie.name)=$($cookie.decrypted_value)"
    } else {
        $finalcookie = $finalcookie + ';' + "$($cookie.name)=$($cookie.decrypted_value)"
    }
}