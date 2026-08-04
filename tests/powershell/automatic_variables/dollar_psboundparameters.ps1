# vybe-test: powershell/automatic_variables/dollar_psboundparameters
function Test-Bound {
    param([string]$Name, [int]$Age)
    return $PSBoundParameters.ContainsKey("Age")
}
$withAge    = Test-Bound -Name "Alice" -Age 30
$withoutAge = Test-Bound -Name "Bob"
if (-not $withAge)   { Write-Host "FAIL: Age should be bound"; exit 1 }
if ($withoutAge)     { Write-Host "FAIL: Age should not be bound"; exit 1 }
Write-Host "PASS"
exit 0
