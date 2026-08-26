# vybe-test: powershell/type_uri_parsing_and_components/is_absolute_uri_flag
$u = [uri]::new("https://example.com", [System.UriKind]::Absolute)
if (-not $u.IsAbsoluteUri) {
    Write-Host "FAIL: IsAbsoluteUri flag check failed"
    exit 1
}
Write-Host "PASS"
exit 0
