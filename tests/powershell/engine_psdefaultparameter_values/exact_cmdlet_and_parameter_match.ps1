# vybe-test: powershell/engine_psdefaultparameter_values/exact_cmdlet_and_parameter_match
$global:PSDefaultParameterValues = @{ "Get-TestWidget:Color" = "NavyBlue" }

function Get-TestWidget {
    [CmdletBinding()]
    param([string]$Color)
    return $Color
}

$res = Get-TestWidget

$global:PSDefaultParameterValues.Clear()

if ($null -eq $res) {
    Write-Host "FAIL: PSDefaultParameterValues exact match did not bind default value"
    exit 1
}

if ($res -ne "NavyBlue") {
    Write-Host "FAIL: expected 'NavyBlue', got '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
