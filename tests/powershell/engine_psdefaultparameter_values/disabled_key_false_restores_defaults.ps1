# vybe-test: powershell/engine_psdefaultparameter_values/disabled_key_false_restores_defaults
$global:PSDefaultParameterValues = @{
    "Get-ConfigState:State" = "Active"
    "Disabled"              = $true
}

function Get-ConfigState {
    [CmdletBinding()]
    param($State)
    return $State
}

# Initially disabled: returns null
$disabledVal = Get-ConfigState
if ($null -ne $disabledVal) {
    Write-Host "FAIL: expected null while Disabled=$true"
    exit 1
}

# Re-enabling by setting Disabled to $false
$global:PSDefaultParameterValues["Disabled"] = $false

$restoredVal = Get-ConfigState

$global:PSDefaultParameterValues.Clear()

if ($restoredVal -ne "Active") {
    Write-Host "FAIL: setting Disabled=$false did not restore default value, got: '$restoredVal'"
    exit 1
}

Write-Host "PASS"
exit 0
