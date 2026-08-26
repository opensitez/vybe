# vybe-test: powershell/type_uri_parsing_and_components/explicit_cast_type_accelerator
$u = [uri]"https://example.com"
if ($u.GetType().Name -ne "Uri") {
    Write-Host "FAIL: [uri] cast failed"
    exit 1
}
Write-Host "PASS"
exit 0
