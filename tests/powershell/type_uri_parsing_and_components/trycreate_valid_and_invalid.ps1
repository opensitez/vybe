# vybe-test: powershell/type_uri_parsing_and_components/trycreate_valid_and_invalid
$outUri = $null
$ok = [uri]::TryCreate("https://valid.org", [System.UriKind]::Absolute, [ref]$outUri)
$bad = [uri]::TryCreate(":::invalid-uri", [System.UriKind]::Absolute, [ref]$outUri)
if (-not $ok -or $bad) {
    Write-Host "FAIL: TryCreate URI validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
