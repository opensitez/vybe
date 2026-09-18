# vybe-test: powershell/engine_psdefaultparameter_values/case_insensitive_key_lookup
# Keys in PSDefaultParameterValues must be matched case-insensitively
$global:PSDefaultParameterValues = @{ "test-casing:myparameter" = "lowercase_key_value" }

function Test-Casing {
    [CmdletBinding()]
    param([string]$MyParameter)
    return $MyParameter
}

$res = Test-Casing

$global:PSDefaultParameterValues.Clear()

if ($res -ne "lowercase_key_value") {
    Write-Host "FAIL: case-insensitive key lookup failed, expected 'lowercase_key_value', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
