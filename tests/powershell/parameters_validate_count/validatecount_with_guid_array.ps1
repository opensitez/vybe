# vybe-test: powershell/parameters_validate_count/validatecount_with_guid_array
function Set-Guids {
    param([ValidateCount(1, 2)][guid[]]$Guids)
    return $Guids.Length
}
$g1 = [guid]::NewGuid()
$res = Set-Guids -Guids $g1
if ($res -ne 1) {
    Write-Host "FAIL: ValidateCount GUID array failed"
    exit 1
}
Write-Host "PASS"
exit 0
