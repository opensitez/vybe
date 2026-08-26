# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/shouldprocess_with_caption_message_query
function Purge-Logs {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Path)
    if ($PSCmdlet.ShouldProcess("Delete all logs at $Path", "Are you sure?", "Purge Logs")) {
        return "Purged"
    }
    return "Aborted"
}
$res = Purge-Logs -Path "/var/log"
if ($res -ne "Purged") {
    Write-Host "FAIL: ShouldProcess with 3 arguments failed"
    exit 1
}
Write-Host "PASS"
exit 0
