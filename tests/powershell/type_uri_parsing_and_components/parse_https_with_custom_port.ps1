# vybe-test: powershell/type_uri_parsing_and_components/parse_https_with_custom_port
$u = [uri]"https://api.example.org:8443/v1/data"
if ($u.Port -ne 8443 -or $u.Scheme -ne "https" -or $u.Host -ne "api.example.org") {
    Write-Host "FAIL: Custom port parsing failed"
    exit 1
}
Write-Host "PASS"
exit 0
