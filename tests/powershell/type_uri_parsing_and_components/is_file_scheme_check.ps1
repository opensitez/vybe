# vybe-test: powershell/type_uri_parsing_and_components/is_file_scheme_check
$u = [uri]"file:///C:/temp/test.txt"
if (-not $u.IsFile -or $u.Scheme -ne "file") {
    Write-Host "FAIL: IsFile check failed"
    exit 1
}
Write-Host "PASS"
exit 0
