# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/confirmimpact_property_high_requires_confirm
function Stop-CriticalService {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="High")]
    param()
}
$cmd = Get-Command Stop-CriticalService
$binding = $cmd.ScriptBlock.Attributes | Where-Object { $_.GetType().Name -eq "CmdletBindingAttribute" }
if ($binding.ConfirmImpact -ne [System.Management.Automation.ConfirmImpact]::High) {
    Write-Host "FAIL: ConfirmImpact High metadata failed"
    exit 1
}
Write-Host "PASS"
exit 0
