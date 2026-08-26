# vybe-test: powershell/parameters_validate_pattern/validatepattern_guid_format
function Set-GuidParam {
    param([ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$GuidStr)
    return $GuidStr
}
$res = Set-GuidParam -GuidStr "12345678-1234-1234-1234-123456789abc"
if ($res -ne "12345678-1234-1234-1234-123456789abc") {
    Write-Host "FAIL: ValidatePattern GUID failed"
    exit 1
}
Write-Host "PASS"
exit 0
