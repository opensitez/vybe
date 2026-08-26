# vybe-test: powershell/pipeline_begin_process_end_blocks/pipeline_dollar_under_inside_process_block
function Print-DollarUnder {
    param([Parameter(ValueFromPipeline=$true)]$Item)
    process { "ITEM:$_" }
}
$res = @("A", "B" | Print-DollarUnder)
if ($res.Length -ne 2 -or $res[0] -ne "ITEM:A" -or $res[1] -ne "ITEM:B") {
    Write-Host "FAIL: `$_ in process block failed, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
