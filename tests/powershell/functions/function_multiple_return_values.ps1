# vybe-test: powershell/functions/function_multiple_return_values
function Get-Stats([int[]]$nums) {
    $min = ($nums | Measure-Object -Minimum).Minimum
    $max = ($nums | Measure-Object -Maximum).Maximum
    $avg = ($nums | Measure-Object -Average).Average
    return $min, $max, $avg
}
$min, $max, $avg = Get-Stats 3, 1, 4, 1, 5, 9, 2, 6
if ($min -ne 1)   { Write-Host "FAIL: min"; exit 1 }
if ($max -ne 9)   { Write-Host "FAIL: max"; exit 1 }
if ($avg -ne 3.875) { Write-Host "FAIL: avg $avg"; exit 1 }
Write-Host "PASS"
exit 0
