# vybe-test: powershell/type_uri_parsing_and_components/uri_equality_case_insensitive_host
$u1 = [uri]"https://EXAMPLE.COM/path"
$u2 = [uri]"https://example.com/path"
if ($u1 -ne $u2) {
    Write-Host "FAIL: URIs with different case in host should compare equal"
    exit 1
}
Write-Host "PASS"
exit 0
