# vybe-test: powershell/parameters_validate_script/validatescript_with_version_major_check
function Check-VersionMajor {
    param([ValidateScript({ $_.Major -ge 7 })][version]$Ver)
    return $Ver.Major
}
$res = Check-VersionMajor -Ver ([version]"7.4.0")
if ($res -ne 7) {
    Write-Host "FAIL: ValidateScript version check failed"
    exit 1
}
Write-Host "PASS"
exit 0
