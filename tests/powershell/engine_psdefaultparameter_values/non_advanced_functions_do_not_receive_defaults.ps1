# vybe-test: powershell/engine_psdefaultparameter_values/non_advanced_functions_do_not_receive_defaults
$global:PSDefaultParameterValues = @{ "Simple-Func:Target" = "DefaultTarget" }

# Simple (non-advanced) function: lacks [CmdletBinding()] and [Parameter()]
function Simple-Func($Target) {
    return $Target
}

# PowerShell specification: simple functions do NOT participate in $PSDefaultParameterValues
$res = Simple-Func

$global:PSDefaultParameterValues.Clear()

if ($null -ne $res) {
    Write-Host "FAIL: simple non-advanced function unexpectedly received PSDefaultParameterValues binding, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
