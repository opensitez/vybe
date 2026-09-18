# vybe-test: powershell/engine_psdefaultparameter_values/caller_argument_overrides_default_value
$global:PSDefaultParameterValues = @{ "Set-TestSetting:Level" = 5 }

function Set-TestSetting {
    [CmdletBinding()]
    param([int]$Level)
    return $Level
}

# Caller passing an explicit argument must override the default table value
$res = Set-TestSetting -Level 99

$global:PSDefaultParameterValues.Clear()

if ($res -ne 99) {
    Write-Host "FAIL: explicit parameter argument did not override default value, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
