# determine framework
if ( $PSVersionTable.PSEdition -eq "Core" ) { $framework = "netcoreapp3.0" }
else { $framework = "net462" }
# determine system architecture
switch -Wildcard ( $env:PROCESSOR_ARCHITECTURE ) {
    "ARM64" { $runtime = "win-arm64" }
    "x86" { $runtime = "win-x86" }
    default { $runtime = "win-x64" }
}
# copy runtime
Join-Path $PSScriptRoot "Microsoft.Web.WebView2.*\runtimes\$runtime\*.dll" -Resolve | ForEach-Object {
    try {
        Copy-Item -Path $_ -Destination (Join-Path $PSScriptRoot "Microsoft.Web.WebView2.*\$framework\" -Resolve) -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Verbose "File is already present, skipping copy: $_"
    }
}
# import assemblies
Join-Path $PSScriptRoot "Microsoft.Web.WebView2.*\$framework\Microsoft.Web.WebView2.*.dll" -Resolve | ForEach-Object {
    Import-Module $_ -ErrorAction Stop
    Write-Verbose "imported assembly $_"
}