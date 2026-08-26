# vybe-test: powershell/parameters_validate_set/validateset_array_parameter_all_valid
function Select-Colors {
    param([ValidateSet("Red", "Green", "Blue")][string[]]$Colors)
    return $Colors.Length
}
$count = Select-Colors -Colors "Red", "Blue"
if ($count -ne 2) {
    Write-Host "FAIL: ValidateSet on array parameter failed, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
