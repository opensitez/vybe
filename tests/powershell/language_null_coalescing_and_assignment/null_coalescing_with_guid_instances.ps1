# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_guid_instances
$g = [guid]::NewGuid()
$nullGuid = $null
$res = $nullGuid ?? $g
if ($res -ne $g) {
    Write-Host "FAIL: ?? with GUID instances failed"
    exit 1
}
Write-Host "PASS"
exit 0
