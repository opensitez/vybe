# vybe-test: powershell/parameters_validate_not_null_or_empty/valid_string_passed
function Set-TargetName {
    param([ValidateNotNullOrEmpty()][string]$Name)
    return "Name:$Name"
}
$res = Set-TargetName -Name "Production"
if ($res -ne "Name:Production") {
    Write-Host "FAIL: ValidateNotNullOrEmpty valid string failed"
    exit 1
}
Write-Host "PASS"
exit 0
