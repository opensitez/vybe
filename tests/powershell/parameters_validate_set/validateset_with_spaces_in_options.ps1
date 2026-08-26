# vybe-test: powershell/parameters_validate_set/validateset_with_spaces_in_options
function Select-Format {
    param([ValidateSet("Plain Text", "Rich Text", "Source Code")][string]$Type)
    return $Type
}
$res = Select-Format -Type "Rich Text"
if ($res -ne "Rich Text") {
    Write-Host "FAIL: ValidateSet with spaces in options failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
