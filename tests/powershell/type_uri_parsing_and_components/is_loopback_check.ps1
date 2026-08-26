# vybe-test: powershell/type_uri_parsing_and_components/is_loopback_check
$u1 = [uri]"http://127.0.0.1:3000"
$u2 = [uri]"http://example.com"
if (-not $u1.IsLoopback -or $u2.IsLoopback) {
    Write-Host "FAIL: IsLoopback check failed"
    exit 1
}
Write-Host "PASS"
exit 0
