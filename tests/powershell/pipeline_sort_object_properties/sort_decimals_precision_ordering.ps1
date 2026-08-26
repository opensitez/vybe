# vybe-test: powershell/pipeline_sort_object_properties/sort_decimals_precision_ordering
[decimal]$d1 = 1.0002
[decimal]$d2 = 1.0001
[decimal]$d3 = 1.0003
$sorted = @($d1, $d2, $d3 | Sort-Object)
if ($sorted[0] -ne $d2 -or $sorted[1] -ne $d1 -or $sorted[2] -ne $d3) {
    Write-Host "FAIL: Sort-Object Decimal precision ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
