# vybe-test: powershell/variables/get_variable_value_parameter
# Get-Variable with -ValueOnly returns the underlying object value directly
$targetVar = "sample string payload"

$retrievedValue = Get-Variable -Name "targetVar" -ValueOnly

if ($retrievedValue -ne "sample string payload") {
    Write-Host "FAIL: unexpected retrieved value: '$retrievedValue'"
    exit 1
}

# ValueOnly should return System.String directly, not PSVariable
if ($retrievedValue.GetType().Name -ne "String") {
    Write-Host "FAIL: expected String type, got $($retrievedValue.GetType().Name)"
    exit 1
}

Write-Host "PASS"
exit 0
