# vybe-test: powershell/type_uri_parsing_and_components/absolute_path_and_query
$u = [uri]"https://example.com/search?q=powershell&page=2"
if ($u.AbsolutePath -ne "/search" -or $u.Query -ne "?q=powershell&page=2") {
    Write-Host "FAIL: Path and query extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
