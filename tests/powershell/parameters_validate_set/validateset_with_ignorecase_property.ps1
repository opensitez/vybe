# vybe-test: powershell/parameters_validate_set/validateset_with_ignorecase_property
function Get-Tag {
    param([ValidateSet("Alpha", "Beta", IgnoreCase=$true)][string]$Tag)
    return $Tag
}
$res = Get-Tag -Tag "alpha"
if ($res -ne "alpha" -and $res -ne "Alpha") {
    Write-Host "FAIL: ValidateSet with IgnoreCase=$true failed"
    exit 1
}
Write-Host "PASS"
exit 0
