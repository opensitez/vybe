# vybe-test: powershell/parameters_validate_range/validaterange_byte_type
function Set-ByteVal {
    param([ValidateRange(10, 200)][byte]$Val)
    return $Val
}
$res = Set-ByteVal -Val 150
if ($res -ne 150) {
    Write-Host "FAIL: ValidateRange byte failed"
    exit 1
}
Write-Host "PASS"
exit 0
