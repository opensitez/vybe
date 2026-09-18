# vybe-test: powershell/engine_psdefaultparameter_values/array_parameter_defaulted
$global:PSDefaultParameterValues = @{
    "Format-Tags:Tags" = @("Alpha", "Beta", "Gamma")
}

function Format-Tags {
    [CmdletBinding()]
    param([string[]]$Tags)
    return ($Tags -join "->")
}

$res = Format-Tags

$global:PSDefaultParameterValues.Clear()

if ($res -ne "Alpha->Beta->Gamma") {
    Write-Host "FAIL: array parameter default failed, got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
