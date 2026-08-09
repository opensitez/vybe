# vybe-test: powershell/should_process/should_process_pipeline_loop
function Clear-Items {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([Parameter(ValueFromPipeline=$true)][int]$Id)
    process {
        if ($PSCmdlet.ShouldProcess("Id:$Id", "Clear")) {
            return "Cleared:$Id"
        }
    }
}
$res = 10..12 | Clear-Items
if ($res[0] -ne "Cleared:10" -or $res[2] -ne "Cleared:12") {
    Write-Host "FAIL: pipeline loop ShouldProcess expected Cleared:10, 11, 12"
    exit 1
}
Write-Host "PASS"
exit 0
