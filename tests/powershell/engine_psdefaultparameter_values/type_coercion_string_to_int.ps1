# vybe-test: powershell/engine_psdefaultparameter_values/type_coercion_string_to_int
$global:PSDefaultParameterValues = @{ "Set-Threshold:Limit" = "500" }

function Set-Threshold {
    [CmdletBinding()]
    param([int]$Limit)
    return $Limit
}

# The string "500" in PSDefaultParameterValues must be coerced to [int]
$res = Set-Threshold

$global:PSDefaultParameterValues.Clear()

if (-not ($res -is [int])) {
    Write-Host "FAIL: default parameter value was not coerced to [int], type is: $($res.GetType().FullName)"
    exit 1
}

if ($res -ne 500) {
    Write-Host "FAIL: expected integer 500, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
