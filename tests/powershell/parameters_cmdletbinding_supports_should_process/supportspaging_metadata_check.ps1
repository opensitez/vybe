# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/supportspaging_metadata_check
function Get-PagedData {
    [CmdletBinding(SupportsPaging=$true)]
    param()
}
$cmd = Get-Command Get-PagedData
if (-not $cmd.Parameters.ContainsKey("First") -or -not $cmd.Parameters.ContainsKey("Skip")) {
    Write-Host "FAIL: SupportsPaging should add First and Skip parameters"
    exit 1
}
Write-Host "PASS"
exit 0
