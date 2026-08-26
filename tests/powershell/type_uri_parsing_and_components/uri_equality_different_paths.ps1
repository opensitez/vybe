# vybe-test: powershell/type_uri_parsing_and_components/uri_equality_different_paths
$u1 = [uri]"https://example.com/path1"
$u2 = [uri]"https://example.com/path2"
if ($u1 -eq $u2) {
    Write-Host "FAIL: URIs with different paths should compare unequal"
    exit 1
}
Write-Host "PASS"
exit 0
