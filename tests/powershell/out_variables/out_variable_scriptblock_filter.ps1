# vybe-test: powershell/out_variables/out_variable_scriptblock_filter
filter Square-Filter {
    param([Parameter(ValueFromPipeline=$true)][int]$n)
    return $n * $n
}
1..3 | Square-Filter -OutVariable sqCap | Out-Null
if ($sqCap[0] -ne 1 -or $sqCap[1] -ne 4 -or $sqCap[2] -ne 9) {
    Write-Host "FAIL: filter cmdlet OutVariable expected 1, 4, 9"
    exit 1
}
Write-Host "PASS"
exit 0
