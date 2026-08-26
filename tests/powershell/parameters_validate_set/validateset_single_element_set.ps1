# vybe-test: powershell/parameters_validate_set/validateset_single_element_set
function Enforce-Single {
    param([ValidateSet("OnlyValid")][string]$Val)
    return $Val
}
$res = Enforce-Single -Val "OnlyValid"
if ($res -ne "OnlyValid") {
    Write-Host "FAIL: Single element ValidateSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
