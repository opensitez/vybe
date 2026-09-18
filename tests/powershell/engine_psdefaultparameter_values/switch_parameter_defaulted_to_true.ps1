# vybe-test: powershell/engine_psdefaultparameter_values/switch_parameter_defaulted_to_true
$global:PSDefaultParameterValues = @{ "Invoke-CleanStep:Force" = $true }

function Invoke-CleanStep {
    [CmdletBinding()]
    param([switch]$Force)
    return $Force.IsPresent
}

$res = Invoke-CleanStep

$global:PSDefaultParameterValues.Clear()

if ($res -ne $true) {
    Write-Host "FAIL: switch parameter defaulted to `$true did not have IsPresent=$true"
    exit 1
}

Write-Host "PASS"
exit 0
