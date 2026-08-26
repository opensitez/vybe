# vybe-test: powershell/parameters_validate_set/validateset_empty_string_allowed_if_in_set
function Check-EmptyAllowed {
    param([ValidateSet("Val1", "")][string]$Val)
    return "OK:$Val"
}
$res = Check-EmptyAllowed -Val ""
if ($res -ne "OK:") {
    Write-Host "FAIL: Empty string in ValidateSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
