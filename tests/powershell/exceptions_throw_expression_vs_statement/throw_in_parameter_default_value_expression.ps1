# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_in_parameter_default_value_expression
function Get-MandatoryParam {
    param([string]$Required = $(throw "Required parameter missing"))
    return "Received:$Required"
}
$r1 = Get-MandatoryParam -Required "Provided"
$caught = $false
try {
    $r2 = Get-MandatoryParam
} catch {
    $caught = $_.Exception.Message.Contains("Required parameter missing")
}
if ($r1 -ne "Received:Provided" -or -not $caught) {
    Write-Host "FAIL: Throw in parameter default value failed"
    exit 1
}
Write-Host "PASS"
exit 0
