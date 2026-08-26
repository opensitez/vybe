# vybe-test: powershell/numeric_random_number_generation/get_random_shuffle_via_cmdlet
$orig = @(1..10)
$shuffled = $orig | Get-Random -Shuffle
if ($shuffled.Count -ne 10 -or ($shuffled | Measure-Object -Sum).Sum -ne 55) {
    Write-Host "FAIL: Get-Random -Shuffle failed"
    exit 1
}
Write-Host "PASS"
exit 0
