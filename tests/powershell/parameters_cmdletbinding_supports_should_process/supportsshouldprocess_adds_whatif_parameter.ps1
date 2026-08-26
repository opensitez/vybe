# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/supportsshouldprocess_adds_whatif_parameter
function Remove-TestFile {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Path)
}
$cmd = Get-Command Remove-TestFile
if (-not $cmd.Parameters.ContainsKey("WhatIf") -or -not $cmd.Parameters.ContainsKey("Confirm")) {
    Write-Host "FAIL: SupportsShouldProcess should automatically add -WhatIf and -Confirm"
    exit 1
}
Write-Host "PASS"
exit 0
