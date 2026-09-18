# vybe-test: powershell/engine_psdefaultparameter_values/dynamic_scriptblock_evaluated_per_call
# A scriptblock in PSDefaultParameterValues is evaluated dynamically on each invocation
$global:PSDefaultParameterValues = @{
    "Get-TimestampedTest:Stamp" = { [System.DateTime]::UtcNow.Year }
}

function Get-TimestampedTest {
    [CmdletBinding()]
    param([int]$Stamp)
    return $Stamp
}

$first = Get-TimestampedTest
$expectedYear = [System.DateTime]::UtcNow.Year

$global:PSDefaultParameterValues.Clear()

if ($first -ne $expectedYear) {
    Write-Host "FAIL: dynamic scriptblock default expected year $expectedYear, got: $first"
    exit 1
}

Write-Host "PASS"
exit 0
