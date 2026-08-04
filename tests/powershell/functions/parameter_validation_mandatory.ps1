# vybe-test: powershell/functions/parameter_validation_mandatory
function Test-Mandatory {
    param(
        [Parameter(Mandatory=$true)]
        $RequiredValue
    )
    return $RequiredValue
}
$result = Test-Mandatory -RequiredValue "test"
if ($result -ne "test") {
    Write-Host "FAIL: expected 'test', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
