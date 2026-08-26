# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/helpuri_metadata_check
function Get-HelpedCmd {
    [CmdletBinding(HelpUri="https://docs.microsoft.com/powershell")]
    param()
}
$cmd = Get-Command Get-HelpedCmd
if ($cmd.HelpUri -ne "https://docs.microsoft.com/powershell") {
    Write-Host "FAIL: HelpUri metadata check failed"
    exit 1
}
Write-Host "PASS"
exit 0
