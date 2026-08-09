# vybe-test: powershell/pipeline_variable/pipeline_variable_scriptblock_filter
filter Multiply-Filter([int]$factor) {
    param([Parameter(ValueFromPipeline=$true)][int]$n)
    return $n * $factor
}
$res = 2..3 | Multiply-Filter 5 -PipelineVariable pv | ForEach-Object { "$pv=>$_" }
if ($res[0] -ne "2=>10" -or $res[1] -ne "3=>15") {
    Write-Host "FAIL: filter cmdlet -PipelineVariable expected 2=>10, 3=>15"
    exit 1
}
Write-Host "PASS"
exit 0
