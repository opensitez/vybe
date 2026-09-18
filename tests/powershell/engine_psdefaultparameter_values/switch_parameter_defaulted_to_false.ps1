# vybe-test: powershell/engine_psdefaultparameter_values/switch_parameter_defaulted_to_false
$global:PSDefaultParameterValues = @{ "Invoke-QuietStep:VerboseOutput" = $false }

function Invoke-QuietStep {
    [CmdletBinding()]
    param([switch]$VerboseOutput)
    return $VerboseOutput.IsPresent
}

$res = Invoke-QuietStep

$global:PSDefaultParameterValues.Clear()

if ($res -ne $false) {
    Write-Host "FAIL: switch parameter defaulted to `$false had IsPresent=$true"
    exit 1
}

Write-Host "PASS"
exit 0
