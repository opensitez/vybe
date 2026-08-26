# vybe-test: powershell/pipeline_begin_process_end_blocks/only_process_block_defined_shorthand
function Simple-ProcOnly {
    process { $_ * 10 }
}
$res = @(1, 2, 3 | Simple-ProcOnly)
if ($res.Length -ne 3 -or $res[0] -ne 10 -or $res[2] -ne 30) {
    Write-Host "FAIL: Process-only shorthand function failed"
    exit 1
}
Write-Host "PASS"
exit 0
