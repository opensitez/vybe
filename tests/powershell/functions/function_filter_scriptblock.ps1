# vybe-test: powershell/functions/function_filter_scriptblock
function Filter-Even {
    param([int[]]$Numbers)
    $Numbers | Where-Object { $_ % 2 -eq 0 }
}
$result = Filter-Even -Numbers @(1, 2, 3, 4, 5, 6)
if ($result.Count -ne 3) {
    Write-Host "FAIL: expected 3 even numbers, got $($result.Count)"
    exit 1
}
if ($result[1] -ne 4) {
    Write-Host "FAIL: expected second even number to be 4"
    exit 1
}
Write-Host "PASS"
exit 0
