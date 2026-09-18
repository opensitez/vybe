# vybe-test: powershell/engine_psdefaultparameter_values/remove_entry_restores_unbound_behavior
$global:PSDefaultParameterValues = @{ "Set-ProfileName:Name" = "TemporaryDefault" }

function Set-ProfileName {
    [CmdletBinding()]
    param($Name)
    return $Name
}

# Bound before removal
$initial = Set-ProfileName
if ($initial -ne "TemporaryDefault") {
    Write-Host "FAIL: initial default failed"
    exit 1
}

# Removing the specific entry
$global:PSDefaultParameterValues.Remove("Set-ProfileName:Name")

# Must return $null now
$unbound = Set-ProfileName

$global:PSDefaultParameterValues.Clear()

if ($null -ne $unbound) {
    Write-Host "FAIL: removing entry did not restore unbound `$null state, got: '$unbound'"
    exit 1
}

Write-Host "PASS"
exit 0
