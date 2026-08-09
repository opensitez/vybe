# vybe-test: powershell/ref_parameters/ref_param_boolean_toggle
function Toggle-Bool([ref]$flag) {
    $flag.Value = -not $flag.Value
}
$state = $false
Toggle-Bool ([ref]$state)
if ($state -ne $true) {
    Write-Host "FAIL: boolean toggle via [ref] expected true"
    exit 1
}
Write-Host "PASS"
exit 0
