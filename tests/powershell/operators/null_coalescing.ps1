# vybe-test: powershell/operators/null_coalescing
$x = $null
$result = $x ?? "default"
if ($result -ne "default") {
    Write-Host "FAIL: expected 'default', got '$result'"
    exit 1
}
$y = "value"
$result2 = $y ?? "default"
if ($result2 -ne "value") {
    Write-Host "FAIL: expected 'value', got '$result2'"
    exit 1
}
Write-Host "PASS"
exit 0
