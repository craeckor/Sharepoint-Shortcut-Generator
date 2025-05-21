[CmdletBinding(SupportsShouldProcess=$true)]
Param()

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
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

try {
    Clear-WebView2Cache -Confirm:$false
} catch {
    return $($_.Exception.Message)
}

# Create a info-Form for informing the user about how to log in into the Microsoft Admincenter
# Create a form
$Micon = [System.Drawing.Icon]::ExtractAssociatedIcon("$workpath\images\icons\Microsoft_logo.ico")
$infoform = New-Object System.Windows.Forms.Form
$infoform.Text = "Microsoft Login Information - Microsoft Admincenter"
$infoform.Size = New-Object System.Drawing.Size(400,325)
$infoform.StartPosition = "CenterScreen"
$infoform.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$infoform.MaximizeBox = $false
$infoform.MinimizeBox = $false
$infoform.ControlBox = $true      # Show close (X) button
$infoform.TopMost = $true         # Always on top
$infoform.Icon = $Micon
$infoform.ShowInTaskbar = $true
$infoform.AutoScalemode = "Dpi"
$infoform.AutoSize = $true

# Add image to the top middle, auto size
$img = New-Object System.Windows.Forms.PictureBox
$img.Image = [System.Drawing.Image]::FromFile("$workpath\images\pictures\important-information-01.png")
$img.SizeMode = "AutoSize"

# Add a red important information label below the image
$label = New-Object System.Windows.Forms.Label
$label.Text = "DO NOT CLOSE the BROWSER WINDOW, it will CLOSE AUTOMATICALLY after the LOGIN.`r`n" +
              "Selecting YES at STAY SIGNED IN is REQUIRED to save the cookies in the browser cache.`r`n" +
              "If you select NO, the cookies will not be saved and the script WILL NOT WORK."
$label.AutoSize = $true
$label.ForeColor = [System.Drawing.Color]::Red

# Set Arial font for label and button
$arialFont = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
$label.Font = $arialFont

# Add a OK button at the bottom middle to close the form and continue
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$okButton.AutoSize = $true
$okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$okButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215) # Windows blue
$okButton.ForeColor = [System.Drawing.Color]::White
$okButton.Padding = '10,5,10,5'
$okButton.FlatAppearance.BorderSize = 0
$okButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$okButton.Add_Click({
    $infoform.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $infoform.Close()
})

$infoform.Controls.Add($img)
$infoform.Controls.Add($label)
$infoform.Controls.Add($okButton)

# Force layout to update control sizes
$infoform.PerformLayout()

[int]$formWidth = $infoform.ClientSize.Width
[int]$imgWidth = $img.Width
[int]$labelWidth = $label.Width
[int]$okButtonWidth = $okButton.Width

$img.Location = New-Object System.Drawing.Point([int](($formWidth - $imgWidth) / 2), 10)
$label.Location = New-Object System.Drawing.Point([int](($formWidth - $labelWidth) / 2), [int]($img.Top + $img.Height + 10))
$okButton.Location = New-Object System.Drawing.Point([int](($formWidth - $okButtonWidth) / 2), [int]($label.Top + $label.Height + 10))

# Show the form and check result
$result = $infoform.ShowDialog()
if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    exit
}

try {
    $response = Invoke-WebView2 -uri "$url" -UrlCloseConditionRegex "$quiturl" -title "Microsoft Login - Microsoft Admincenter"
    if ($response -eq $null) {
        return "Error: No response from WebView2"
    }
    if ($response -notlike "*https://admin.microsoft.com*") {
        Write-Error -Message "Login failed. Please check your credentials and try again."
        return "Error: Login failed. Please check your credentials and try again."
    }
} catch {
    return $($_.Exception.Message)
}

try {
    $cookies = Get-DecryptedCookiesInfo -Browser "CustomChrome" -DomainName "$searchdomain" -customPath "$cookiePath"
    
    # Check if required cookies exist by looking at the name property of each cookie object
    $hasOIDCAuthCookie = $false
    $hasRootAuthToken = $false
    
    foreach ($cookie in $cookies) {
        if ($cookie.name -eq "OIDCAuthCookie") {
            $hasOIDCAuthCookie = $true
        }
        if ($cookie.name -eq "RootAuthToken") {
            $hasRootAuthToken = $true
        }
    }
    
    if (-not ($hasOIDCAuthCookie -and $hasRootAuthToken)) {
        Write-Error -Message "Required cookies not found. Please check your login status."
        return "Error: Required cookies not found. Please check your login status."
    }
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