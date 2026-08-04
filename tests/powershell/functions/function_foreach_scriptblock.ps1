# vybe-test: powershell/functions/function_foreach_scriptblock
function Double-Array {
    param([int[]]$Numbers)
    $Numbers | ForEach-Object { $_ * 2 }
}
$result = Double-Array -Numbers @(1, 2, 3)
if ($result.Count -ne 3) {
    Write-Host "FAIL: expected 3 results, got $($result.Count)"
    exit 1
}
if ($result[2] -ne 6) {
    Write-Host "FAIL: expected last result to be 6"
    exit 1
}
Write-Host "PASS"
exit 0
