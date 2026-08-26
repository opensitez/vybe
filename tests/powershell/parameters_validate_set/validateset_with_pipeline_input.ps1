# vybe-test: powershell/parameters_validate_set/validateset_with_pipeline_input
function Filter-Protocol {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [ValidateSet("HTTP", "HTTPS", "FTP")]
        [string]$Proto
    )
    process { "OK:$Proto" }
}
$res = "HTTPS" | Filter-Protocol
if ($res -ne "OK:HTTPS") {
    Write-Host "FAIL: Pipeline input to ValidateSet failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
