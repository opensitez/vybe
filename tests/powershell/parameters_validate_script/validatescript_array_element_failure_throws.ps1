# vybe-test: powershell/parameters_validate_script/validatescript_array_element_failure_throws
function Check-AllPositive2 {
    param([ValidateScript({ $_ -gt 0 })][int[]]$Nums)
    return $Nums
}
$caught = $false
try {
    $x = Check-AllPositive2 -Nums 1, -2, 3
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when one array item fails ValidateScript"
    exit 1
}
Write-Host "PASS"
exit 0
