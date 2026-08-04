# vybe-test: powershell/arrays/foreach_statement
$sum = 0
foreach ($item in @(10, 20, 30)) {
    $sum += $item
}
if ($sum -ne 60) {
    Write-Host "FAIL: expected 60, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
