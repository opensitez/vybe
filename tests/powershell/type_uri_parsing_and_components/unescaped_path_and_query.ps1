# vybe-test: powershell/type_uri_parsing_and_components/unescaped_path_and_query
$u = [uri]"https://example.com/path%20with%20spaces/item?tag=c%23"
if ($u.AbsolutePath -ne "/path%20with%20spaces/item") {
    Write-Host "FAIL: AbsolutePath escaping failed"
    exit 1
}
Write-Host "PASS"
exit 0
