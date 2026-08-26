# vybe-test: powershell/numeric_random_number_generation/get_random_with_count_returns_subset
$arr = @(1, 2, 3, 4, 5, 6)
$sample = $arr | Get-Random -Count 3
if ($sample.Count -ne 3) {
    Write-Host "FAIL: Get-Random -Count 3 failed, got count $($sample.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
