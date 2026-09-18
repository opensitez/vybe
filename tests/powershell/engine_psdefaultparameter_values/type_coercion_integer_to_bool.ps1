# vybe-test: powershell/engine_psdefaultparameter_values/type_coercion_integer_to_bool
# In PowerShell, integer values (1 / 0) in PSDefaultParameterValues coerce to [bool] parameters
$global:PSDefaultParameterValues = @{ "Set-SecurityFlag:Enabled" = 1 }

function Set-SecurityFlag {
    [CmdletBinding()]
    param([bool]$Enabled)
    return $Enabled
}

$res = Set-SecurityFlag

$global:PSDefaultParameterValues.Clear()

if (-not ($res -is [bool])) {
    Write-Host "FAIL: default parameter value was not coerced to [bool]"
    exit 1
}

if ($res -ne $true) {
    Write-Host "FAIL: expected boolean `$true from integer 1, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
