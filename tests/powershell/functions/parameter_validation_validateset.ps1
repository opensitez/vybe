# vybe-test: powershell/functions/parameter_validation_validateset
function Test-ValidateSet {
    param(
        [ValidateSet("Small", "Medium", "Large")]
        $Size
    )
    return $Size
}
$result = Test-ValidateSet -Size "Medium"
if ($result -ne "Medium") {
    Write-Host "FAIL: expected 'Medium', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
