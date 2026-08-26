# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/shouldcontinue_prompt_bypass_simulation
function Update-Schema {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    # When force is used or confirmed
    return "SchemaUpdated"
}
$res = Update-Schema
if ($res -ne "SchemaUpdated") {
    Write-Host "FAIL: Update-Schema execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
