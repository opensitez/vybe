# vybe-test: powershell/engine_psdefaultparameter_values/multiple_parameters_defaulted_on_same_function
$global:PSDefaultParameterValues = @{
    "Initialize-Connection:TargetHost" = "127.0.0.1"
    "Initialize-Connection:Port"       = 9000
    "Initialize-Connection:UseSSL"     = $true
}

function Initialize-Connection {
    [CmdletBinding()]
    param(
        [string]$TargetHost,
        [int]$Port,
        [bool]$UseSSL
    )
    return "$TargetHost`:$Port`:SSL=$UseSSL"
}

$res = Initialize-Connection

$global:PSDefaultParameterValues.Clear()

if ($res -ne "127.0.0.1:9000:SSL=True") {
    Write-Host "FAIL: multiple parameter defaults failed, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
