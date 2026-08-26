# vybe-test: powershell/parameters_validate_count/validatecount_large_max_count
function Set-Batch {
    param([ValidateCount(1, 1000)][int[]]$Batch)
    return $Batch.Length
}
$arr = @(1..50)
$res = Set-Batch -Batch $arr
if ($res -ne 50) {
    Write-Host "FAIL: ValidateCount large max failed"
    exit 1
}
Write-Host "PASS"
exit 0
