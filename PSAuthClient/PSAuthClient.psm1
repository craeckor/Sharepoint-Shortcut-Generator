function Clear-WebView2Cache {
    <#
    .SYNOPSIS
    Deletes the WebView2 cache folder.
    .DESCRIPTION
    Removes PSAuthClient WebView2 user data folder (UDF) which is used to store browser data such as cookies, permissions and cached resources.
    .EXAMPLE
    Clear-WebView2Cache
    Deletes the WebView2 cache folder.
    #>
    [cmdletbinding(SupportsShouldProcess, ConfirmImpact = 'High')]param()
    if ($PSCmdlet.ShouldProcess("PSAuthClientWebview2Cache", "delete")) { 
        if ( (test-path "$env:temp\PSAuthClientWebview2Cache\") ) { Remove-Item "$env:temp\PSAuthClientWebview2Cache\" -Recurse -Force }
    }
}
function ConvertFrom-JsonWebToken {
    <#
    .SYNOPSIS
    Convert (decode) a JSON Web Token (JWT) to a PowerShell object.

    .DESCRIPTION
    Convert (decode) a JSON Web Token (JWT) to a PowerShell object.

    .PARAMETER jwtInput
    The JSON Web Token (string) to be must be in the form of a valid JWT.

    .EXAMPLE
    PS> ConvertFrom-JsonWebToken "ew0KICAidHlwIjogIkpXVCIsDQogICJhbGciOiAiUlMyNTYiDQp9.ew0KICAiZXhwIjogMTcwNjc4NDkyOSwNCiAgImVjaG8..."
    header    : @{typ=JWT; alg=RS256}
    exp       : 1706784929
    echo      : Hello World!
    nbf       : 1706784629
    sub       : PSAuthClient
    iss       : https://example.org
    jti       : 27913c80-40d1-46a3-89d5-d3fb9f0d1e4e
    iat       : 1706784629
    aud       : PSAuthClient
    signature : OHIxRGxuaXVLTjh4eXhRZ0VWYmZ3SHNlQ29iOUFBUVRMK1dqWUpWMEVXMD0

    #>
    param ( 
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [ValidatePattern("^e[yJ|w0]([a-zA-Z0-9_-]+[.]){2}", Options = "None")]
        [string]$jwtInput
    )
    $response = New-Object -TypeName PSObject
    [pscustomobject]$jwt = $jwtInput -split "[.]"
    for ( $i = 0; $i -lt $jwt.count; $i++ ) {
        try { $data = ConvertFrom-Base64UrlEncoding $jwt[$i] | ConvertFrom-Json }
        catch { $data = $jwt[$i] }
        switch ( $i ) {
            0 { $response | Add-Member -NotePropertyName header -TypeName NoteProperty $data }
            1 { $response | Add-Member -NotePropertyName payload -TypeName NoteProperty $data  }
            2 { $response | Add-Member -NotePropertyName signature -TypeName NoteProperty $data  }
        }
    }
    # ...We prefer ordered output
    return $response | Select-Object header, signature -ExpandProperty payload | Select-Object header, * -ErrorAction SilentlyContinue
}
function Get-OidcDiscoveryMetadata {
    <#
    .SYNOPSIS
    Retreive OpenID Connect Discovery endpoint metadata
    .DESCRIPTION
    Retreive OpenID Connect Discovery endpoint metadata.
    .PARAMETER uri
    The URI of the OpenID Connect Discovery endpoint.
    .EXAMPLE
    PS> Get-OidcDiscoveryMetadata "https://example.org"
    Attempts to retreive OpenID Connect Discovery endpoint metadata from 'https://example.org/.well-known/openid-configuration'.
    .EXAMPLE
    PS> Get-OidcDiscoveryMetadata "https://login.microsoftonline.com/common"
    token_endpoint                        : https://login.microsoftonline.com/common/oauth2/token
    token_endpoint_auth_methods_supported : {client_secret_post, private_key_jwt, client_secret...}
    jwks_uri                              : https://login.microsoftonline.com/common/discovery/keys
    response_modes_supported              : {query, fragment, form_post}
    subject_types_supported               : {pairwise}
    id_token_signing_alg_values_supported : {RS256}
    response_types_supported              : {code, id_token, code id_token, token id_tokenÔÇª}
    scopes_supported                      : {openid}
    issuer                                : https://sts.windows.net/{tenantid}/
    microsoft_multi_refresh_token         : True
    authorization_endpoint                : https://login.microsoftonline.com/common/oauth2/auth...
    device_authorization_endpoint         : https://login.microsoftonline.com/common/oauth2/devi...
    http_logout_supported                 : True
    frontchannel_logout_supported         : True
    end_session_endpoint                  : https://login.microsoftonline.com/common/oauth2/logo...
    claims_supported                      : {sub, iss, cloud_instance_name, cloud_instance_host...}
    check_session_iframe                  : https://login.microsoftonline.com/common/oauth2/chec...
    userinfo_endpoint                     : https://login.microsoftonline.com/common/openid/user...
    kerberos_endpoint                     : https://login.microsoftonline.com/common/kerberos
    tenant_region_scope                   : 
    cloud_instance_name                   : microsoftonline.com
    cloud_graph_host_name                 : graph.windows.net
    msgraph_host                          : graph.microsoft.com
    rbac_url                              : https://pas.windows.net
    #>
    Param($uri)
    if ( $uri -notmatch "^http(s)?://" ) { $uri = "https://$uri" }
    if ( $uri -notmatch "[.]well-known" ) { $uri = "$uri/.well-known/openid-configuration" }
    return Invoke-RestMethod $uri -Method GET -Verbose:$false
}
function ConvertFrom-Base64UrlEncoding {
    <#
    .SYNOPSIS
    Convert a a base64url string to text.
    .DESCRIPTION
    Base64-URL-encoding is a minor variation on the typical Base64 encoding method, but uses URL-safe characters instead. 
    .PARAMETER value
    The base64url encoded string to convert.
    .PARAMETER rawBytes
    Return output as byteArray instead of the string.
    .EXAMPLE
    ConvertFrom-Base64UrlEncoding -value "eyJ0eXAiOiJKV1QiLCJhbGciOiJ..."
    {"typ":"JWT","alg":"RS256","kid":"kWbkha6qs8wsTnBwiNYOhHbnAw"}
    #>
    param( 
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [string]$value, 
        [Parameter (Mandatory = $false)]
        [switch]$rawBytes 
    )
    # change - back to +, and _ back to /
    $value = $value -replace "-","+" -replace "_","/"
    # determine padding
    switch ( $value.Length % 4 ) {
        0 { break }
        2 { $value += '==' }
        3 { $value += '=' }
    }
    try { 
        $byteArray = [system.convert]::FromBase64String($value)
        if ( $rawBytes ) { return $byteArray }
        else { return [System.Text.Encoding]::Default.GetString($byteArray) }
    }
    catch { throw $_ }
}
function ConvertTo-Base64UrlEncoding {
    <# 
    .SYNOPSIS
    Convert a a string to base64url encoding.
    .DESCRIPTION
    Base64-URL-encoding is a minor variation on the typical Base64 encoding method, but uses URL-safe characters instead. 
    .PARAMETER value
    string to convert
    .EXAMPLE
    ConvertTo-Base64UrlEncoding -value "{"typ":"JWT","alg":"RS25....}
    eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6ImtXYmthYTZxczh3c1RuQndpaU5ZT2hIYm5BdyJ9
    #>
    param ([Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]$value)
    if ( $value.GetType().Name -eq "String" ) { $value = [System.Text.Encoding]::UTF8.GetBytes($value) }
    # change + to -, and / to _, then trim the trailing = from the end.
    return ( [System.Convert]::ToBase64String($value) -replace '\+','-' -replace '/','_' -replace '=' )
}
function Get-RandomString { 
    <#
    .SYNOPSIS
    Generate a random string.
    .DESCRIPTION
    Generate a random string of a given length, by default between 43 and 128 characters.
    .PARAMETER Length
    The length of the string to generate.
    .EXAMPLE
    Get-RandomString
    95TFIttXFdwvhW8DCXVhlHqsld62U_NlxlQe.YqdN.Hm5xs8S3.bTISQ
    #>
    param( [int]$Length = ((43..128) | get-random) )
    $charArray = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-._~".ToCharArray()
    if ( $PSVersionTable.PSVersion -ge "7.4.0" ) { $randomizedCharArray = for ( $i = 0; $i -lt $Length; $i++ ) { $charArray[(0..($charArray.count-1) | Get-SecureRandom)] } }
    else { $randomizedCharArray = for ( $i = 0; $i -lt $Length; $i++ ) { $charArray[(0..($charArray.count-1) | Get-Random)] } }
    return [string](-join $randomizedCharArray)
}
function Get-UnixTime {
    <#
    .SYNOPSIS
    Get the current time in Unix time.
    .DESCRIPTION
    Get the time in Unix time (seconds since the Epoch)
    .PARAMETER date
    The date to convert to Unix time. If not specified, the current date and time is used.
    .EXAMPLE
    Get-UnixTime
    1706961267
    #>
    param( [Parameter(Mandatory=$false,ValueFromPipeline=$true)] [DateTime]$date = (Get-Date) )
    return [int64](New-TimeSpan -Start (Get-Date "1970-01-01T00:00:00Z").ToUniversalTime() -End ($date).ToUniversalTime()).TotalSeconds    
}
function Invoke-WebView2 {
    <#
    .SYNOPSIS
    PowerShell Interactive browser window using WebView2.

    .DESCRIPTION
    Uses WebView2 (a embedded edge browser) to allow embeded web technologies (HTML, CSS and JavaScript).

    .PARAMETER uri
    The URL to browse.

    .PARAMETER UrlCloseConditionRegex
    (optional, default:'error=') - The form will close when the URL matches the regex.

    .PARAMETER failoverToWindowsFormsWebBrowser
    (optional, defalt:true) - If WebView2 fails to initialize, the form will close and the script will continue using Windows.Forms.WebBrowser (IE11).

    .PARAMETER allowSingleSignOnUsingOSPrimaryAccount
    (optional, default:true) Determines whether to enable Single Sign-on with Azure Active Directory (AAD) resources inside WebView using the logged in Windows account.

    .PARAMETER title
    (optional, default:'PowerShell WebView') - Window-title of the form.

    .PARAMETER Width
    (optional, default:600) - Width of the form.

    .PARAMETER Height
    (optional, default:800) - Height of the form.

    .PARAMETER userAgent
    (optional) Only supported for WebView2, will be ignored if failing over to Windows.Forms.WebBrowser.

    .EXAMPLE
    PS> Invoke-WebView2 -uri "https://microsoft.com/devicelogin" -UrlCloseConditionRegex "//appverify$" -title "Device Code Flow" | Out-Null

    Starts a form with a WebView2 control, and navigates to https://microsoft.com/devicelogin. The form will close when the URL matches the regex '//appverify$'.

    #>
    param(
        [parameter(Position = 0, Mandatory = $true, HelpMessage="The URL to browse.")]
        [string]$uri,

        [parameter( Mandatory = $false, HelpMessage="Form close condition by regex (URL)")]
        [string]$UrlCloseConditionRegex = "error=[^&]*",

        [parameter( Mandatory = $false, HelpMessage="WebView2 failover to System.Windows.Forms.WebBrowser (IE11)")]
        [bool]$failoverToWindowsFormsWebBrowser = $true,

        [parameter( Mandatory = $false, HelpMessage="msSingleSignOnOSForPrimaryAccountIsShared")]
        [bool]$allowSingleSignOnUsingOSPrimaryAccount = $true,

        [parameter( Mandatory = $false, HelpMessage="Forms window title")]
        [string]$title = "PowerShell WebView",

        [parameter( Mandatory = $false, HelpMessage="Forms window width")]
        [int]$Width = "600",

        [parameter( Mandatory = $false, HelpMessage="Forms window height")]
        [int]$Height = "800",

        [parameter( Mandatory = $false, HelpMessage="Customize the User-Agent presented in the HTTP Header.")]
        $userAgent
    
    )
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.web.webview2.core.corewebview2environmentoptions.allowsinglesignonusingosprimaryaccount?view=webview2-dotnet-1.0.2210.55
    if ( $allowSingleSignOnUsingOSPrimaryAccount ) { $env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--enable-features=msSingleSignOnOSForPrimaryAccountIsShared" }
    else { $env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = $null }
    # initialize WebView2
    try { 
        $web = New-Object 'Microsoft.Web.WebView2.WinForms.WebView2'
        $web.CreationProperties = New-Object 'Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties'
        $web.CreationProperties.UserDataFolder = "$env:temp\PSAuthClientWebview2Cache\" 
        $web.Dock = "Fill"
        $web.source = $uri
        if ( $userAgent ) { $web.add_CoreWebView2InitializationCompleted({$web.CoreWebView2.Settings.UserAgent = $userAgent}) }
        # close form on completion (match redirectUri) navigation
        $web.Add_SourceChanged( {
            if ( $web.source.AbsoluteUri -match $UrlCloseConditionRegex )  { $Form.close() | Out-Null }
        })
    }
    # if WebView2 fails to initialize, try to use Windows.Forms.WebBrowser
    catch {
        if ( $failoverToWindowsFormsWebBrowser ) { 
            Write-Warning "Failed to initialize WebView2, trying to use Windows.Forms.WebBrowser."
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width = $Width; Height = $Height; Url = $uri }
            # Close form on completion (match redirectUri) navigation
            $docCompletedEvent = {
                if ( $web.Url.AbsoluteUri -match $UrlCloseConditionRegex )  { $form.Close() }
            }
            $web.Add_DocumentCompleted($docCompletedEvent)
            $web.ScriptErrorsSuppressed = $true
            $title = $title + " [COMPATABILITY MODE]"
    }
        else { throw $_ }
    }
    # Create form
    $form = New-Object System.Windows.Forms.Form -Property @{Width=$Width;Height=$Height;Text=$title} -ErrorAction Stop
    # Add the WebBrowser control to the form
    $form.Controls.Add($web)
    $form.Add_Shown( { $form.Activate() } )
    $form.ShowDialog() | Out-Null
    $response = $web.Source
    $web.Dispose()
    return $response
}
function New-HttpListener {     
    <#
    .SYNOPSIS
    Create new HttpListener object and listen for incoming requests.
    .DESCRIPTION
    Create new HttpListener object and listen for incoming requests.
    .PARAMETER prefix
    The URI prefix handled by the HttpListener object, typically a redirect_uri.
    .EXAMPLE
    New-HttpListener -prefix "http://localhost:8080/"
    Waits for a request on http://localhost:8080/ and returns the post data.
    .EXAMPLE
    $job = Start-Job -ScriptBlock (Get-Command 'New-HttpListener').ScriptBlock -ArgumentList $redirect_uri
    Starts a Job which waits for a request on $redirect_uri.
    #>
    param(
        [parameter(Position = 0, Mandatory = $true, HelpMessage="The URI prefix handled by the HttpListener object, typically the redirect_uri.")]
        [string]$prefix
    )
    try {
        # start http listener
        $httpListener = New-Object System.Net.HttpListener
        $httpListener.Prefixes.Add($prefix)
        $httpListener.Start()
        # wait for request
        $context = $httpListener.GetContext()
        # read post request
        $form_post = [System.IO.StreamReader]::new($context.Request.InputStream).ReadToEnd()
        # send a response
        $context.Response.StatusCode = 200
        $context.Response.ContentType = 'application/json'
        $responseBytes = [System.Text.Encoding]::UTF8.GetBytes('')
        $context.Response.OutputStream.Write($responseBytes, 0, $responseBytes.Length)
        # clean up
        $context.Response.Close()
        $httpListener.Close()
    }
    catch { throw $_ } 
    return $form_post
}
function Invoke-OAuth2AuthorizationEndpoint { 
    <#
    .SYNOPSIS
    Interact with a OAuth2 Authorization endpoint.

    .DESCRIPTION
    Uses WebView2 (embedded browser based on Microsoft Edge) to request authorization, this ensures support for modern web pages and capabilities like SSO, Windows Hello, FIDO key login, etc

    OIDC and OAUTH2 Grants such as Authorization Code Flow with PKCE, Implicit Flow (including form_post) and Hybrid Flows are supported

    .PARAMETER uri
    Authorization endpoint URL.

    .PARAMETER client_id
    The identifier of the client at the authorization server.

    .PARAMETER redirect_uri
    The client callback URI for the authorization response.

    .PARAMETER response_type
    Tells the authorization server which grant to execute. Default is code.

    .PARAMETER scope
    One or more space-separated strings indicating which permissions the application is requesting. 

    .PARAMETER usePkce
    Proof Key for Code Exchange (PKCE) improves security for public clients by preventing and authorization code attacks. Default is $true.

    .PARAMETER response_mode
    OPTIONAL Informs the Authorization Server of the mechanism to be used for returning Authorization Response. Determined by response_type, if not specified.

    .PARAMETER customParameters
    Hashtable with custom parameters added to the request uri (e.g. domain_hint, prompt, etc.) both the key and value will be url encoded. Provided with state, nonce or PKCE keys these values are used in the request (otherwise values are generated accordingly).

    .PARAMETER userAgent
    OPTIONAL Custom User-Agent string to be used in the WebView2 browser.

    .EXAMPLE
    PS> Invoke-OAuth2AuthorizationEndpoint -uri "https://acc.spotify.com/authorize" -client_id "2svXwWbFXj" -scope "user-read-currently-playing" -redirect_uri "http://localhost"
    code_verifier                  xNTKRgsEy_u2Y.PQZTmUbccYd~gp7-5v4HxS7HVKSD2fE.uW_yu77HuA-_sOQ...
    redirect_uri                   https://localhost
    client_id                      2svXwWbFXj
    code                           AQDTWHSP6e3Hx5cuJh_85r_3m-s5IINEcQZzjAZKdV4DP_QRqSHJzK_iNB_hN...

    A request for user authorization is sent to the /authorize endpoint along with a code_challenge, code_challenge_method and state param. 
    If successful, the authorization server will redirect back to the redirect_uri with a code which can be exchanged for an access token.
    
    .EXAMPLE
    PS> Invoke-OAuth2AuthorizationEndpoint -uri "https://example.org/oauth2/authorize" -client_id "0325" -redirect_uri "http://localhost" -scope "user.read" -response_type "token" -usePkce:$false -customParameters @{ login = "none" }
    expires_in                     4146
    expiry_datetime                01.02.2024 10:56:06
    scope                          User.Read profile openid email
    session_state                  5c044a21-543e-4cbc-a94r-d411ddec5a87
    access_token                   eyJ0eXAiQiJKV1QiLCJub25jZSI6InAxYTlHksH6bktYdjhud3VwMklEOGtUM...
    token_type                     Bearer

    Implicit Grant, will return a access_token if successful.
    #>
    [Alias('Invoke-AuthorizationEndpoint','authorize')]
    [OutputType([hashtable])]
    param(
        [parameter(Position = 0, Mandatory = $true)]
        [string]$uri,

        [parameter( Mandatory = $true)]
        [string]$client_id,

        [parameter( Mandatory = $false)]
        [string]$redirect_uri,
        
        [parameter( Mandatory = $false)]
        [validatePattern("(code)?(id_token)?(token)?(none)?")]
        [string]$response_type = "code",

        [parameter( Mandatory = $false)]
        [string]$scope,

        [parameter( Mandatory = $false)]
        [bool]$usePkce = $true,

        [parameter( Mandatory = $false)]
        [ValidateSet("query","fragment","form_post")]
        [string]$response_mode,

        [parameter( Mandatory = $false)]
        [hashtable]$customParameters,

        [parameter( Mandatory = $false)]
        [string]$userAgent
    )

    # Determine which protocol is being used.
    if ( $response_type -eq "token" -or ($response_type -match "^code$" -and $scope -notmatch "openid" ) ) { $protocol = "OAUTH"; $nonce = $null }
    else { $protocol = "OIDC"
        # ensure scope contains openid for oidc flows
        if ( $scope -notmatch "openid" ) { Write-Warning "Invoke-OAuth2AuthorizationRequest: Added openid scope to request (OpenID requirement)."; $scope += " openid" }
        # ensure nonce is present for id_token validation
        if ( $customParameters -and $customParameters.Keys -match "^nonce$" ) { [string]$nonce = $customParameters["nonce"] }
        else { [string]$nonce = Get-RandomString -Length ( (32..64) | get-random ) }
    }

    # state for CSRF protection (optional, but recommended)
    if ( $customParameters -and $customParameters.Keys -match "^state$" ) { [string]$state = $customParameters["state"] }
    else { [string]$state = Get-RandomString -Length ( (16..21) | get-random ) }
    
    # building the request uri
    Add-Type -AssemblyName System.Web -ErrorAction Stop
    $uri += "?response_type=$($response_type)&client_id=$([System.Web.HttpUtility]::UrlEncode($client_id))&state=$([System.Web.HttpUtility]::UrlEncode($state))"
    if ( $redirect_uri ) { $uri += "&redirect_uri=$([System.Web.HttpUtility]::UrlEncode($redirect_uri))" } 
    if ( $scope ) { $uri += "&scope=$([System.Web.HttpUtility]::UrlEncode($scope))" }
    if ( $nonce ) { $uri += "&nonce=$([System.Web.HttpUtility]::UrlEncode($nonce))" }
    
    # PKCE for code flows
    if ( $response_type -notmatch "code" -and $usePkce ) { write-verbose "Invoke-OAuth2AuthorizationRequest: PKCE is not supported for implicit flows." }
    else { 
        if ( $usePkce ) {
            # pkce provided in custom parameters
            if ( $customParameters -and $customParameters.Keys -match "^code_challenge$" ) {
                $pkce = @{ code_challenge = $customParameters["code_challenge"] }
                if ( $customParameters.Keys -match "^code_challenge_method$" ) { $pkce.code_challenge_method = $customParameters["code_challenge_method"] }
                else { Write-Warning "Invoke-OAuth2AuthorizationRequest: code_challenge_method not specified, defaulting to 'S256'."; $pkce.code_challenge_method = "S256" }
                if ( $customParameters.Keys -match "^code_verifier$" ) { $pkce.code_verifier = $customParameters["code_verifier"] }
            }
            # generate new pkce challenge
            else { $pkce = New-PkceChallenge }
            # add to request uri
            $uri += "&code_challenge=$($pkce.code_challenge)&code_challenge_method=$($pkce.code_challenge_method)"
        }
    }

    # Add custom parameters to request uri
    if ( $customParameters ) { 
        foreach ( $key in ($customParameters.Keys | Where-Object { $_ -notmatch "^nonce$|^state$|^code_(challenge(_method)?$|verifier)$" }) ) { 
            $urlEncodedValue = [System.Web.HttpUtility]::UrlEncode($customParameters[$key])
            $urlEncodedKey = [System.Web.HttpUtility]::UrlEncode($key)
            $uri += "&$urlEncodedKey=$urlEncodedValue"
        }
    }
    Write-Verbose "Invoke-OAuth2AuthorizationRequest $protocol request uri $uri"
    
    # if form_post: start http.sys listener as job and give it some time to start
    if ( $response_mode -eq "form_post" ) { 
        $uri += "&response_mode=form_post"
        $job = Start-Job -ScriptBlock (Get-Command 'New-HttpListener').ScriptBlock -ArgumentList $redirect_uri
        Start-Sleep -Milliseconds 500
    }

    # authorization request (interactive)
    $webViewParams = @{
        uri = $uri
        title = "Authorization code flow"
        UrlCloseConditionRegex = "$($redirect_uri)?.*(?:code=([^&]+)|error=([^&]+))|^$redirect_uri"
    }
    if ( $userAgent ) { $webViewParams.userAgent = $userAgent }
    $webSource = Invoke-WebView2 @webViewParams
    
    # if form post - retreive job (post) after interaction has been complete
    if ( $response_mode -eq "form_post" ) {
        Write-Verbose "Invoke-OAuth2AuthorizationRequest: http.sys waiting for form_post request. (timeout 10s)"
        Wait-Job $job -Timeout 10 | Out-Null
        if ( $job.State -eq "Running" ) { 
            Write-Verbose "Invoke-OAuth2AuthorizationRequest: http.sys did not receive form_post request before timeout (10s)."; Stop-Job $job; Remove-Job $job
            throw "Invoke-OAuth2AuthorizationRequest: did not receive form_post request before timeout (10s)."
        }
        try { $jobData = Receive-Job $job -ErrorAction Stop }
        catch { 
            if ( $_.Exception.Message -match "Access is denied." ) { throw "Unabled to start http.sys listener, please run as admin." }
            else { throw "http.sys listener failed, error: $($jobData.Exception.Message)" }
        }
        finally { Remove-Job $job }
        Write-Verbose "Invoke-OAuth2AuthorizationRequest: http.sys form_post request received."
        $webSource = @{ Fragment = $jobData; Query = $null }
    }

    # When the window closes (WebView2), the script will continue and retreive the depending on the response_mode and content.
    if( $webSource.query -match "code=" ) {
        $response = @{}
        $response.code = [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['code']
        $response.state = [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['state']
        if ( $protocol -eq "oidc" ) { 
            $response.nonce = $nonce 
            if ( [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['access_token'] ) { $response.access_token = [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['access_token'] }
            if ( [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['id_token'] ) { $response.access_token = [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['id_token'] }
        }
    } 
    elseif ( $webSource.Query -match "error=" ) { 
        $errorDetails = [ordered]@{}
        $errorDetails.error = [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['error']
        $errorDetails.error_uri = [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['error_uri']
        $errorDetails.error_description = [System.Web.HttpUtility]::ParseQueryString($webSource.Query)['error_description']
        throw ($errorDetails | convertTo-Json)
    }
    elseif ( $webSource.Fragment -match "token|error=" ) {
        $response = @{}
        foreach ( $item in (($webSource.Fragment -split "#|&") | Where-Object { $_ -ne "" }) ) { 
            $key = $item.split("=")[0]
            $value = $item.split("=")[1]
            $value = [System.Web.HttpUtility]::UrlDecode($value)
            $response.$key = $value
        }
        if ( $protocol -eq "oidc" ) { $response.nonce = $nonce }
        if ( $webSource.Fragment -match "error=" ) { throw ($response | Select-Object * -ExcludeProperty state | convertTo-Json) }
    }
    else { throw "invalid response received" }

    # if code grant, add client_id, code_verifier and redirect_uri to output (needed for token exchange) 
    if ( $response_type -match "code" ) {
        $response.client_id = $client_id
        if ( $usePkce -and $pkce.Keys -match "^code_verifier$" ) { $response.code_verifier = $pkce.code_verifier }
        if ( $redirect_uri ) { $response.redirect_uri = $redirect_uri }
    }

    # verify state
    if ( $response.state -ne $state ) { throw "State mismatch!`nExpected '$state', got '$($response.state)'." }
    else { Write-Verbose "Invoke-OAuth2AuthorizationRequest: Validated state in response." }

    # if OIDC - validate id_token nonce
    if ( $nonce -and $response.ContainsKey("id_token") ) { 
        $decodedToken = ConvertFrom-JsonWebToken $response.id_token -Verbose:$false
        if ( $decodedToken.nonce -ne $nonce ) { throw "id_token nonce mismatch`nExpected '$nonce', got '$($decodedToken.nonce)'." }
        Write-Verbose "Invoke-OAuth2TokenExchange: Validated nonce in id_token response."
    }

    # add expiry_datetime to output
    if ( $response.ContainsKey("expires_in") ) { 
        $response["expires_in"] = [int]$response["expires_in"]
        $response.expiry_datetime = (get-date).AddSeconds($response.expires_in) 
    }

    # remove state from output
    if ( $response.ContainsKey("state") ) { $response.Remove("state") }

    # badabing badaboom
    return $response
}
function Invoke-OAuth2DeviceAuthorizationEndpoint {
    <#
    .SYNOPSIS
    OAuth2.0 Device Authorization Endpoint Interaction
    .DESCRIPTION
    Get a unique device verification code and an end-user code from the device authorization endpoint, which then can be used to request tokens from the token endpoint.
    .PARAMETER uri
    The URI of the OAuth2.0 Device Authorization endpoint.
    .PARAMETER client_id
    The client identifier.
    .PARAMETER scope
    The scope of the access request.
    .EXAMPLE
    PS> Invoke-OAuth2DeviceAuthorizationEndpoint -uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode" -client_id $splat.client_id -scope $splat.scope
    user_code        : L8EFTXRY3
    device_code      : LAQABAAEAAAAmoFfGtYxvRrNriQdPKIZ-2b64dTFbGcmRF3rSBagHQGtBcyz0K_XV8ltq-nXz...
    verification_uri : https://microsoft.com/devicelogin
    expires_in       : 900
    interval         : 5
    message          : To sign in, use a web browser to open the page https://microsoft.com/devi...
    #>
    [Alias('Invoke-DeviceAuthorizationEndpoint','deviceauth')]
    param (
        [parameter( Position = 0, Mandatory = $true, HelpMessage="Token endpoint URL.")]
        [string]$uri,
        [parameter( Mandatory = $true )]
        [string]$client_id,
        [parameter( Mandatory = $false )]
        [string]$scope
    )
    $requestBody = @{}
    $requestBody.client_id = $client_id
    if ( $scope ) { $requestBody.scope = $scope}
    $response = Invoke-RestMethod -Uri $uri -Body $requestBody
    Write-Warning -Message $response.message -ErrorAction SilentlyContinue 
    return $response
}
function Invoke-OAuth2TokenEndpoint { 
    <#
    .SYNOPSIS
    Token Exchange by OAuth2.0 Token Endpoint interaction.

    .DESCRIPTION
    Forge and send token exchange requests to the OAuth2.0 Token Endpoint.

    .PARAMETER uri
    Authorization endpoint URL.

    .PARAMETER client_id
    The identifier of the client at the authorisation server. (required if no other client authentication is present)

    .PARAMETER redirect_uri
    The client callback URI for the response. (required if it was included in the initial authorization request)

    .PARAMETER scope
    One or more space-separated strings indicating which permissions the application is requesting. 

    .PARAMETER code
    Authorization code received from the authorization server.

    .PARAMETER code_verifier
    Code_verifier, required if code_challenge was used in the authorization request (PKCE).

    .PARAMETER device_code
    Device verification code received from the authorization server.

    .PARAMETER client_secret
    client credential as string or securestring

    .PARAMETER client_auth_method
    OPTIONAL client (secret) authentication method. (default: client_secret_post)

    .PARAMETER client_certificate
    client credential as x509certificate2, RSA Private key or cert location 'Cert:\CurrentUser\My\THUMBPRINT'

    .PARAMETER nonce
    OPTIONAL nonce value used to associate a Client session with an ID Token, and to mitigate replay attacks.

    .PARAMETER refresh_token
    Refresh token received from the authorization server.

    .PARAMETER customHeaders
    Hashtable with custom headers to be added to the request uri (e.g. User-Agent, Origin, Referer, etc.).

    .EXAMPLE
    PS> $code = Invoke-OAuth2AuthorizationEndpoint -uri $authorization_endpoint @splat
    PS> Invoke-OAuth2TokenEndpoint -uri $token_endpoint @code
    token_type      : Bearer
    scope           : User.Read profile openid email
    expires_in      : 4948
    ext_expires_in  : 4948
    access_token    : eyJ0eXAiOiJKV1QiLCJujHu5ZSI6IllQUWFERGtkVEczUjJIYm5tWFlhOG1QbHk1ZTlwckJybm...
    refresh_token   : 0.AUcAjvFfm8BTokWLwpwMkH65xiGBP5hz2ZpErJuc3chlhOUNAVw.AgABAAEAAAAmoFfGtYxv...
    id_token        : eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6ImtXYmthYTZxczh3c1RuQndpaU5ZT2...
    expiry_datetime : 01.02.2024 12:51:33

    Invokes the Token Endpoint by passing parameters from the authorization endpoint response such as code, code_verifier, nonce, etc.

    .EXAMPLE 
    PS> Invoke-OAuth2TokenEndpoint -uri $token_endpoint -scope ".default" -client_id "123" -client_secret "secret"
    token_type      : Bearer
    expires_in      : 3599
    ext_expires_in  : 3599
    access_token    : eyJ0eXAiOiJKV1QikjY6b25jZSI6IkZ4YTZ4QmloQklGZjFPT0FqQZQ4LTl5WUEtQnpqdXNzTn...
    expiry_datetime : 01.02.2024 12:33:01

    Client authentication using client_secret_post

    .EXAMPLE
    PS> Invoke-OAuth2TokenEndpoint -uri $token_endpoint -scope ".default" -client_id "123" "Cert:\CurrentUser\My\8kge399dddc5521e04e34ac19fe8f8759ba021b8"
    token_type      : Bearer
    expires_in      : 3599
    ext_expires_in  : 3599
    access_token    : eyJ0eXYiOiJKV1QiLCJusk6jZSI6IlNkYU9lOTdtY0NqS0g1VnRURjhTY3JSMEgwQ0hje24fR1...
    expiry_datetime : 01.02.2024 12:36:17

    Client authentication using private_key_jwt

    #>
    [Alias('Invoke-TokenEndpoint','token')]
    [cmdletbinding(DefaultParameterSetName='code')]
    param(
        [parameter( Position = 0, Mandatory = $true, ParameterSetName='client_certificate')]    
        [parameter( Position = 0, Mandatory = $true, ParameterSetName='client_secret')]
        [parameter( Position = 0, Mandatory = $true, ParameterSetName='code')]
        [parameter( Position = 0, Mandatory = $true, ParameterSetName='device_code')]
        [parameter( Position = 0, Mandatory = $true, ParameterSetName='refresh')]
        [string]$uri,

        [parameter( Mandatory = $false, ParameterSetName='client_certificate')]
        [parameter( Mandatory = $false, ParameterSetName='client_secret')]
        [parameter( Mandatory = $false, ParameterSetName='code')]
        [parameter( Mandatory = $false, ParameterSetName='device_code')]
        [parameter( Mandatory = $false, ParameterSetName='refresh')]
        [string]$client_id,

        [parameter( Mandatory = $false, ParameterSetName='client_certificate')]
        [parameter( Mandatory = $false, ParameterSetName='client_secret')]
        [parameter( Mandatory = $false, ParameterSetName='code')]
        [parameter( Mandatory = $false, ParameterSetName='refresh')]
        [string]$redirect_uri,

        [parameter(Mandatory = $false, ParameterSetName='client_certificate')]
        [parameter(Mandatory = $false, ParameterSetName='client_secret')]
        [parameter(Mandatory = $false, ParameterSetName='code')]
        [parameter(Mandatory = $false, ParameterSetName='refresh')]
        [string]$scope,

        [parameter( Mandatory = $false, ParameterSetName='client_certificate')]
        [parameter( Mandatory = $false, ParameterSetName='client_secret')]
        [parameter( Mandatory = $false, ParameterSetName='code')]
        [string]$code,

        [parameter( Mandatory = $false, ParameterSetName='client_certificate')]
        [parameter( Mandatory = $false, ParameterSetName='client_secret')]
        [parameter( Mandatory = $false, ParameterSetName='code')]
        [string]$code_verifier,

        [parameter( Mandatory = $true, ParameterSetName='device_code')]
        [string]$device_code,

        [parameter( Mandatory = $true, ParameterSetName='client_secret')]
        [parameter( Mandatory = $false, ParameterSetName='code')]
        [parameter( Mandatory = $false, ParameterSetName='refresh')]
        $client_secret,

        [parameter( Mandatory = $false, ParameterSetName='client_secret', HelpMessage="client_credential type")]
        [parameter( Mandatory = $false, ParameterSetName='refresh')]
        [ValidateSet('client_secret_basic','client_secret_post','client_secret_jwt')]
        $client_auth_method = "client_secret_post",

        [parameter( Mandatory = $true, ParameterSetName='client_certificate', HelpMessage="private_key_jwt")]
        [parameter( Mandatory = $false, ParameterSetName='refresh', HelpMessage="private_key_jwt")]
        $client_certificate,

        [parameter( Mandatory = $false, ParameterSetName='client_certificate')]
        [parameter( Mandatory = $false, ParameterSetName='client_secret')]
        [parameter( Mandatory = $false, ParameterSetName='code')]
        [parameter( Mandatory = $false, ParameterSetName='refresh')]
        [string]$nonce,

        [parameter( Mandatory = $true, ParameterSetName='refresh')]
        $refresh_token,

        [parameter( Mandatory = $false, ParameterSetName='client_certificate')]
        [parameter( Mandatory = $false, ParameterSetName='client_secret')]
        [parameter( Mandatory = $false, ParameterSetName='code')]
        [parameter( Mandatory = $false, ParameterSetName='device_code')]
        [parameter( Mandatory = $false, ParameterSetName='refresh')]
        [hashtable]$customHeaders
    )

    $payload = @{}
    $payload.headers = @{ 'Content-Type' = 'application/x-www-form-urlencoded' }
    $payload.method  = 'Post'
    $payload.uri     =  $uri

    # Custom headers provided
    if ( $customHeaders ) { 
        if ( $customHeaders.Keys -contains "Content-Type" ) { $payload.headers = $customHeaders }
        else { $payload.headers += $customHeaders }
    }
    
    # Build request body and determine grant_type
    $requestBody = @{}
    if ( $client_id ) { $requestBody.client_id = $client_id }
    if ( $redirect_uri ) { $requestBody.redirect_uri = $redirect_uri }
    if ( $scope ) { $requestBody.scope = $scope }
    if ( ( $client_secret -or $client_certificate) -and !$code -and !$refresh_token ) { 
        $requestBody.grant_type = "client_credentials" 
    }
    elseif ( $code ) { 
        $requestBody.grant_type = "authorization_code"
        $requestBody.code = $code 
        if ( $code_verifier ) { $requestBody.code_verifier = $code_verifier }
    }
    elseif ( $device_code ) { 
        $requestBody.grant_type = "urn:ietf:params:oauth:grant-type:device_code"
        $requestBody.device_code = $device_code
    }
    elseif ( $refresh_token ) { 
        $requestBody.grant_type = "refresh_token"
        $requestBody.refresh_token = $refresh_token
    }

    # client authentication
    if ( $client_secret ) { 
        if ( $client_secret.GetType().Name -eq "SecureString" ) { 
            # Psv5 ConvertFrom-SecureString does not have -AsPlainText Param
            if ( $PSVersionTable.PSEdition -eq "Core" ) { $client_secret = $client_secret | ConvertFrom-SecureString -AsPlainText }
            else {
                # secureString.toBinaryString.toStringUnit
                $client_secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($client_secret)
                )
            }
        }
        switch ( $client_auth_method ) {
            "client_secret_basic" { $payload.headers.Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$($client_id):$($client_secret)")) }
            "client_secret_post" { $requestBody.client_secret = $client_secret }
            "client_secret_jwt" { $requestBody.client_assertion = (New-Oauth2JwtAssertion -aud $uri -iss $client_id -sub $client_id -client_secret $client_secret -ErrorAction Stop).client_assertion_jwt }
        }
    }
    elseif ( $client_certificate) {
        if ( $client_certificate.GetType().Name -notmatch "^X509Certificate|^RSA" ) {
            try { $client_certificate = Get-Item $client_certificate -ErrorAction Stop }
            catch { throw $_ }
        }
        $requestBody.client_assertion = (New-Oauth2JwtAssertion -aud $uri -iss $client_id -sub $client_id -client_certificate $client_certificate -ErrorAction Stop).client_assertion_jwt
    }
    if ( $client_certificate -or $client_auth_method -eq "client_secret_jwt" ) { $requestBody.client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"}

    $payload.body = $requestBody
    write-verbose ($payload | ConvertTo-Json -Compress)
    try { $response = Invoke-RestMethod @payload -Verbose:$false }
    catch { throw $_ }

    # validate id_token nonce
    if ( $nonce -and $response.id_token ) { 
        $decodedToken = ConvertFrom-JsonWebToken $response.id_token -Verbose:$false
        if ( $decodedToken.nonce -ne $nonce ) { throw "id_token nonce mismatch`nExpected '$nonce', got '$($decodedToken.nonce)'." }
        Write-Verbose "Invoke-OAuth2TokenExchange: Validated nonce in id_token"
    }

    # add expiry datetime
    if ( $response.expires_in ) { $response | Add-Member -NotePropertyName expiry_datetime -TypeName NoteProperty (get-date).AddSeconds($response.expires_in) }
    
    # badabing badaboom
    return $response
}
function New-Oauth2JwtAssertion {
    <#
    .SYNOPSIS
    Create a JWT Assertion for OAuth2.0 Client Authentication.
    
    .DESCRIPTION
    Create a JWT Assertion for OAuth2.0 Client Authentication.

    .PARAMETER issuer
    iss, must contain the client_id of the OAuth Client.

    .PARAMETER subject
    sub, must contain the client_id of the OAuth Client.

    .PARAMETER audience
    aud, should be the URL of the Authorization Server's Token Endpoint.

    .PARAMETER jwtId
    jti, unique token identifier. Random GUID by default.

    .PARAMETER customClaims
    Hashtable with custom claims to be added to the JWT payload (assertion).

    .PARAMETER client_certificate
    Location Cert:\CurrentUser\My\THUMBPRINT, x509certificate2 or RSA Private key.

    .PARAMETER key_id
    kid, key identifier for assertion header

    .PARAMETER client_secret
    clientsecret for HMAC signature

    .EXAMPLE
    PS> New-Oauth2JwtAssertion -issuer $client_id -subject $client_id -audience $oidcDiscoveryMetadata.token_endpoint -client_certificate $cert

    client_assertion_jwt          ew0KICAidHlwIjogIkpXVCIsDQogICJhbGciOiAiUlMyNTYiDQp9.ew0KICAia...
    client_assertion_type          urn:ietf:params:oauth:client-assertion-type:jwt-bearer
    header                         @{typ=JWT; alg=RS256}
    payload                        @{iss=PSAuthClient; nbf=1706785754; iat=1706785754; sub=PSAu...}

    #>
    [cmdletbinding(DefaultParameterSetName='private_key_jwt')]
    param(
        [parameter( Position = 0, Mandatory = $true, ParameterSetName='private_key_jwt', HelpMessage="iss, must contain the client_id of the OAuth Client.")]
        [parameter( Position = 0, Mandatory = $true, ParameterSetName='client_secret_jwt', HelpMessage="iss, must contain the client_id of the OAuth Client.")]
        [string]$issuer,

        [parameter( Position = 1, Mandatory = $true, ParameterSetName='private_key_jwt', HelpMessage="sub, must contain the client_id of the OAuth Client.")]
        [parameter( Position = 1, Mandatory = $true, ParameterSetName='client_secret_jwt', HelpMessage="sub, must contain the client_id of the OAuth Client.")]
        [string]$subject,

        [parameter( Position = 2, Mandatory = $true, ParameterSetName='private_key_jwt', HelpMessage="aud, should be the URL of the Authorization Server's Token Endpoint.")]
        [parameter( Position = 2, Mandatory = $true, ParameterSetName='client_secret_jwt', HelpMessage="aud, should be the URL of the Authorization Server's Token Endpoint.")]
        [string]$audience,

        [parameter( Position = 3, Mandatory = $false, ParameterSetName='private_key_jwt', HelpMessage="jti, unique token identifier.")]
        [parameter( Position = 3, Mandatory = $false, ParameterSetName='client_secret_jwt', HelpMessage="jti, unique token identifier.")]
        [string]$jwtId = [string]([guid]::NewGuid()),

        [parameter( Mandatory = $false, ParameterSetName='private_key_jwt', HelpMessage="Hashtable with custom claims.")]
        [parameter( Mandatory = $false, ParameterSetName='client_secret_jwt', HelpMessage="Hashtable with custom claims.")]
        [hashtable]$customClaims,

        [parameter( Mandatory = $true, ParameterSetName='private_key_jwt', HelpMessage="Location Cert:\CurrentUser\My\THUMBPRINT, x509certificate2 or RSA Private key.")]
        $client_certificate,

        [parameter( Mandatory = $false, ParameterSetName='private_key_jwt', HelpMessage="key identifier for assertion header")]
        $key_id,

        [parameter( Mandatory = $true, ParameterSetName='client_secret_jwt', HelpMessage="ClientSecret")]
        $client_secret
    )
    $jwtHeader = @{ alg = "RS256"; typ = "JWT" }
    # certificate properties
    if ( $client_certificate ) {
        if ( $client_certificate.GetType().Name -notmatch "^X509Certificate|^RSA" ) {
            try { $client_certificate = Get-Item $client_certificate -ErrorAction Stop }
            catch { throw $_ }
        }
        if ( $client_certificate.GetType().Name -match "^X509Certificate" ) { $jwtHeader.x5t = ConvertTo-Base64urlencoding $client_certificate.GetCertHash() }
        if ( $key_id ) { $jwtHeader.kid = $key_id }
    }
    elseif ( $client_secret ) {
        if ( $client_secret.GetType().Name -eq "SecureString" ) { 
            # Psv5 ConvertFrom-SecureString does not have -AsPlainText Param
            if ( $PSVersionTable.PSEdition -eq "Core" ) { $client_secret = $client_secret | ConvertFrom-SecureString -AsPlainText }
            else {
                # secureString.toBinaryString.toStringUnit
                $client_secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($client_secret)
                )
            }
        }
    }
    $jwtHeader = $jwtHeader | ConvertTo-Json
    # build assertion payload
    $jwtClaims = @{
        aud = $audience                 # URL of the resource using the JWT to authenticate to
        exp = (Get-UnixTime) + 300      # expiration time of the token
        jti = $jwtId                    # (optional) unique identifier for the token
        iat = Get-UnixTime              # (optional) time the token was issued
        iss = $issuer                   # issuer of token (client_id)
        sub = $subject                  # subject of the token (client_id)
        nbf = Get-UnixTime              # (optional) time before which the token is not valid
    } 
    if ( $customClaims ) { foreach ( $key in $customClaims.Keys ) { $jwtClaims.$key = $customClaims[$key] } }
    $jwtClaims = $jwtClaims | ConvertTo-Json
    # unsigned assertion in base64url encoding
    $jwtAssertion = (ConvertTo-Base64urlencoding $jwtHeader) + "." + (ConvertTo-Base64urlencoding $jwtClaims)
    # assertion signing - cert or secret
    if ( $client_certificate ) {
        if ( $client_certificate.GetType().Name -match "^X509" ) { $signature = convertTo-Base64urlencoding $client_certificate.PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($jwtAssertion),[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1) }
        elseif ( $client_certificate.GetType().Name -match "^RSA" ) { $signature = convertTo-Base64urlencoding $client_certificate.SignData([System.Text.Encoding]::UTF8.GetBytes($jwtAssertion),[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1) }
    }
    else { 
        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.key = [Text.Encoding]::UTF8.GetBytes($client_secret)
        $signature = $hmacsha.ComputeHash([Text.Encoding]::UTF8.GetBytes($jwtAssertion))
        $signature = convertTo-Base64urlencoding ( [Convert]::ToBase64String($signature) )
    }
    # Finalize response
    $response = [ordered]@{}
    $response.client_assertion_jwt = $jwtAssertion + "." + $Signature
    $response.client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
    $response.header = $jwtHeader | ConvertFrom-Json
    $response.payload = $jwtClaims  | ConvertFrom-Json
    return $response
}
function New-PkceChallenge { 
    <#
    .SYNOPSIS
    Generate code_verifier and code_challenge for PKCE (authorization code flow).
    .DESCRIPTION
    Generate code_verifier and code_challenge for PKCE (authorization code flow).
    .EXAMPLE
    PS> New-PkceChallenge
    code_verifier                  Vpq2YXOsD~1DRM-jBPR6bt8R-3dWQAHNLVLUIDxh7SkWpOT3A0grpenqKne5rAHcVKsTi-ya8-lGBxJ0NS7zavdcFbfdN0yFQ5kYOFbWBh3
    code_challenge                 TW-3r-6mxRWjhkkxmYOabLlwIQ0JkQ0ndxzOSLJvCoU
    code_challenge_method          S256
    #>
    # Generate code_verifier and code_challenge for PKCE (authorization code flow).
    $response = [ordered]@{}
    # code_verifier (should be a random string using the characters "[A-Z,a-z,0-9],-._~" between 43 and 128 characters long)
    [string]$response.code_verifier = Get-RandomString
    # dereive code_challenge from code_verifier (SHA256 hash) and encode using base64urlencoding (fallback to plain)
    try { 
        $hashAlgorithm = [System.Security.Cryptography.HashAlgorithm]::Create('sha256') 
        $hashInBytes = $hashAlgorithm.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($response.code_verifier))
        [string]$response.code_challenge = ConvertTo-Base64urlencoding $hashInBytes
        [string]$response.code_challenge_method = "S256"
    }
    catch {
        [string]$response.code_challenge = $response.code_verifier
        [string]$response.code_challenge_method = "plain"
    }  
    return $response
}
function Test-JsonWebTokenSignature {
    <#
    .SYNOPSIS
    Test the signature of a JSON Web Token (JWT)

    .DESCRIPTION
    Automatically attempt to test the signature of a JSON Web Token (JWT) by using the issuer discovery metadata to get the signing certificate if no signing certificate or secret was provided.

    .PARAMETER jwtInput
    The JSON Web Token (string) to be must be in the form of a valid JWT.

    .PARAMETER SigningCertificate
    X509Certificate2 object to be used for RSA signature verification.

    .PARAMETER client_secret
    Client secret to be used for HMAC signature verification.

    .EXAMPLE
    PS> Test-JsonWebTokenSignature -jwtInput $jwt

    Decodes the JWT and attempts to verify the signature using the issuer discovery metadata to get the signing certificate if no signing certificate or secret was provided.

    .EXAMPLE
    PS> Test-JsonWebTokenSignature -jwtInput $jwt -SigningCertificate $cert

    Decodes the JWT and attempts to verify the signature using the provided certificate.

    .EXAMPLE
    PS> Test-JsonWebTokenSignature -jwtInput $jwt -client_secret $secret

    Decodes the JWT and attempts to verify the signature using the provided client secret.

    #>
    [OutputType([bool])]
    [cmdletbinding(DefaultParameterSetName='certificate')]
    param ( 
        [Parameter( Position = 0, Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='certificate')]
        [Parameter( Position = 0, Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='client_secret')]
        [ValidatePattern("^e[yJ|w0]([a-zA-Z0-9_-]+[.]){2}", Options = "None")]
        $jwtInput,

        [Parameter(Mandatory = $false, ParameterSetName='certificate')]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$SigningCertificate,

        [Parameter(Mandatory = $true, ParameterSetName='client_secret')]
        $client_secret
    )

    try { $decodedJwt = ConvertFrom-JsonWebToken $jwtInput }
    catch { throw $_ }
    
    # attempt to get signing key from discovery metadata if not specified
    if ( !$signingCertificate -and !$client_secret ) { 
        try {
            Write-Verbose "Test-JWTSignature: Attempting to get signing key from issuer discovery metadata"
            $signingKey = (Invoke-RestMethod -uri (get-oidcDiscoveryMetadata $decodedJwt.iss).jwks_uri -Verbose:$false).keys | Where-Object kid -eq $decodedJwt.header.kid
            if ( !$signingKey ) { throw "Test-JWTSignature: Unable to get signing key from issuer discovery metadata, please specify certificate in input parameter." }
            $signingCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new( [convert]::FromBase64String( $signingKey.x5c ) )
            Write-Verbose "Test-JWTSignature: Retrieved keyId $($signingKey.kid) from issuer $($decodedJwt.iss)"
        } 
        catch { throw $_ }
    }

    # data to be verified
    $data = [System.Text.Encoding]::UTF8.GetBytes( ($jwtInput -split "[.]")[0..1] -join "." ) # (YES, DOT-included ,_. ffs.)

    if ( $SigningCertificate ) { 
        # public key
        if ( $decodedJwt.header.alg -match "^[RS|PS]" ) { 
            $publicKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($signingCertificate) 
            if ( $decodedJwt.header.alg -match "^PS" ) { $padding = [Security.Cryptography.RSASignaturePadding]::Pss }
            else { $padding = [Security.Cryptography.RSASignaturePadding]::Pkcs1 }
        }
        elseif ( $decodedJwt.header.alg -match "^ES" ) { $publicKey = [System.Security.Cryptography.X509Certificates.ECDsaCertificateExtensions]::GetECDsaPublicKey($signingCertificate) }
        else { throw "Test-JWTSignature: Unsupported algorithm $($decodedJwt.header.alg)" } # https://www.iana.org/assignments/jose/jose.xhtml#IESG

        # signature
        [byte[]]$sig = ConvertFrom-Base64UrlEncoding $decodedJwt.signature -rawBytes

        # alg
        $alg = "SHA$(($decodedJwt.header.alg -replace '[a-z]'))"
        Write-Verbose "Test-JWTSignature: attempting to verify $($decodedJwt.header.alg) signature"
        if ( $padding ) { [bool]$response = $publicKey.VerifyData( $data, $sig, $alg, [Security.Cryptography.RSASignaturePadding]::Pkcs1) }
        else { [bool]$response = $publicKey.VerifyData( $data, $sig, $alg ) }
    }
    elseif ( $client_secret ) {
        Write-Verbose "Test-JWTSignature: attempting to verify HMAC signature"
        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.key = [Text.Encoding]::UTF8.GetBytes($client_secret)
        $signature = $hmacsha.ComputeHash($data)
        $signature = convertTo-Base64urlencoding ( [Convert]::ToBase64String($signature) )
        if ( $signature -eq $decodedJwt.signature ) { $response = $true }
        else { $response = $false }
    }
    return $response
}
