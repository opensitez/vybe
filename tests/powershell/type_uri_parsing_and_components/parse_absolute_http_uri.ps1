# vybe-test: powershell/type_uri_parsing_and_components/parse_absolute_http_uri
$u = [uri]"http://example.com/path/file.html"
if ($u.Scheme -ne "http" -or $u.Host -ne "example.com" -or $u.AbsolutePath -ne "/path/file.html") {
    Write-Host "FAIL: Absolute HTTP URI parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
