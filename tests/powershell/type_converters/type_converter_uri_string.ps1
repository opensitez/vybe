# vybe-test: powershell/type_converters/type_converter_uri_string
$u = [uri]"http://localhost:80"
if ($u.Port -ne 80 -or $u.Host -ne "localhost") {
    Write-Host "FAIL: string to [uri] conversion expected localhost:80"
    exit 1
}
Write-Host "PASS"
exit 0
