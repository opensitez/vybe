# vybe-test: powershell/script_parameters/alias_parameter
function Test-ScriptParam {
    [CmdletBinding()]
    param([string]$Name = "DefaultVal")
    return $Name
}
$res = Test-ScriptParam -Name "CustomVal"
if ($res -eq "CustomVal") {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
