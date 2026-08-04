# vybe-test: powershell/advanced_functions/advanced_function_parameter_validation
function Test-Advanced {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    return $Name
}
$result = Test-Advanced -Name 'value'
if ($result -ne 'value') {
    Write-Host "FAIL: expected value, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
