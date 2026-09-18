# vybe-test: powershell/engine_psdefaultparameter_values/caller_explicit_null_overrides_default
$global:PSDefaultParameterValues = @{ "Invoke-TestTask:Tag" = "DefaultTag" }

function Invoke-TestTask {
    [CmdletBinding()]
    param($Tag = "Fallback")
    return $Tag
}

# Passing $null explicitly must be honored, overriding the PSDefaultParameterValues entry
$res = Invoke-TestTask -Tag $null

$global:PSDefaultParameterValues.Clear()

if ($null -ne $res) {
    Write-Host "FAIL: explicit `$null argument did not override default parameter value, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
