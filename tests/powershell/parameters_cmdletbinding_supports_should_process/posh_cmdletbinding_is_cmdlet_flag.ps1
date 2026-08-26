# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/posh_cmdletbinding_is_cmdlet_flag
function Check-CmdletBindingActive {
    [CmdletBinding()]
    param()
    return ($PSCmdlet -ne $null)
}
$res = Check-CmdletBindingActive
if ($res -ne $true) {
    Write-Host "FAIL: `$PSCmdlet automatic variable must be populated in [CmdletBinding()]"
    exit 1
}
Write-Host "PASS"
exit 0
