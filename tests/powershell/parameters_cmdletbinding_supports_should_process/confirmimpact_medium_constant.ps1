# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/confirmimpact_medium_constant
function Medium-ImpactCmd {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="Medium")]
    param()
}
$cmd = Get-Command Medium-ImpactCmd
$binding = $cmd.ScriptBlock.Attributes | Where-Object { $_.GetType().Name -eq "CmdletBindingAttribute" }
if ($binding.ConfirmImpact -ne [System.Management.Automation.ConfirmImpact]::Medium) {
    Write-Host "FAIL: ConfirmImpact Medium metadata check failed"
    exit 1
}
Write-Host "PASS"
exit 0
