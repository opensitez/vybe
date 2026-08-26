# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/confirmimpact_low_constant
function Safe-ReadCmd {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="Low")]
    param()
}
$cmd = Get-Command Safe-ReadCmd
$binding = $cmd.ScriptBlock.Attributes | Where-Object { $_.GetType().Name -eq "CmdletBindingAttribute" }
if ($binding.ConfirmImpact -ne [System.Management.Automation.ConfirmImpact]::Low) {
    Write-Host "FAIL: ConfirmImpact Low metadata check failed"
    exit 1
}
Write-Host "PASS"
exit 0
