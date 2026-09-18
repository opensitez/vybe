# vybe-test: powershell/engine_psdefaultparameter_values/disabled_key_temporarily_suppresses_defaults
$global:PSDefaultParameterValues = @{
    "Get-ConfigItem:Key" = "DatabaseURL"
}

function Get-ConfigItem {
    [CmdletBinding()]
    param($Key)
    return $Key
}

# Verify it works initially
$initial = Get-ConfigItem
if ($initial -ne "DatabaseURL") {
    Write-Host "FAIL: initial default binding failed"
    exit 1
}

# The 'Disabled' key toggles off all default bindings without clearing the table
$global:PSDefaultParameterValues["Disabled"] = $true

$suppressed = Get-ConfigItem

$global:PSDefaultParameterValues.Clear()

if ($null -ne $suppressed) {
    Write-Host "FAIL: setting Disabled=$true did not suppress default binding, got: '$suppressed'"
    exit 1
}

Write-Host "PASS"
exit 0
