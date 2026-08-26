# vybe-test: powershell/numeric_random_number_generation/shuffle_array_in_place
$rnd = [System.Random]::new(99)
[int[]]$arr = @(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
$rnd.Shuffle($arr)
if ($arr.Length -ne 10 -or ($arr | Measure-Object -Sum).Sum -ne 55) {
    Write-Host "FAIL: Random Shuffle corrupted elements"
    exit 1
}
Write-Host "PASS"
exit 0
