# vybe-test: powershell/type_uri_parsing_and_components/to_string_canonical
$u = [uri]"HTTP://example.com:80/index.html"
if ($u.ToString() -ne "http://example.com/index.html") {
    Write-Host "FAIL: ToString canonical form failed"
    exit 1
}
Write-Host "PASS"
exit 0
