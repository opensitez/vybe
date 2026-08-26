# vybe-test: powershell/error_record_category_info/error_record_in_pipeline_non_terminating_stream
function Emit-Mixed {
    [CmdletBinding()]
    param()
    1
    Write-Error "Non-terminating error" -Category ReadError
    2
}
$out = @(Emit-Mixed 2>$null)
if ($out.Length -ne 2 -or $out[0] -ne 1 -or $out[1] -ne 2) {
    Write-Host "FAIL: Non-terminating error stream separation failed"
    exit 1
}
Write-Host "PASS"
exit 0
