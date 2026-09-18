# vybe-test: powershell/engine_psdefaultparameter_values/hashtable_parameter_defaulted
$global:PSDefaultParameterValues = @{
    "Connect-Service:Options" = @{ Timeout = 30; Retries = 3 }
}

function Connect-Service {
    [CmdletBinding()]
    param([hashtable]$Options)
    return "$($Options.Timeout):$($Options.Retries)"
}

$res = Connect-Service

$global:PSDefaultParameterValues.Clear()

if ($res -ne "30:3") {
    Write-Host "FAIL: hashtable parameter default failed, expected '30:3', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
