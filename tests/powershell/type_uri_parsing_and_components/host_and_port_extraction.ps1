# vybe-test: powershell/type_uri_parsing_and_components/host_and_port_extraction
$u = [uri]"http://localhost:8080"
if ($u.Host -ne "localhost" -or $u.Port -ne 8080) {
    Write-Host "FAIL: Host and port extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
