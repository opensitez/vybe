# vybe-test: powershell/arrays/array_length_vs_count
$arr = @(1, 2, 3, 4, 5)
if ($arr.Length -ne 5) {
    Write-Host "FAIL: expected Length 5, got $($arr.Length)"
    exit 1
}
if ($arr.Count -ne 5) {
    Write-Host "FAIL: expected Count 5, got $($arr.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
