# vybe-test: powershell/engine_psdefaultparameter_values/clear_method_removes_all_defaults
$global:PSDefaultParameterValues = @{
    "Get-Alpha:Val" = "AlphaVal"
    "Get-Beta:Val"  = "BetaVal"
}

function Get-Alpha { [CmdletBinding()] param($Val) $Val }
function Get-Beta  { [CmdletBinding()] param($Val) $Val }

# Verify both are initially defaulted
if ((Get-Alpha) -ne "AlphaVal" -or (Get-Beta) -ne "BetaVal") {
    Write-Host "FAIL: initial defaults failed"
    exit 1
}

# Clear all defaults
$global:PSDefaultParameterValues.Clear()

$a = Get-Alpha
$b = Get-Beta

if ($null -ne $a -or $null -ne $b) {
    Write-Host "FAIL: .Clear() did not remove all default values, got a='$a', b='$b'"
    exit 1
}

Write-Host "PASS"
exit 0
