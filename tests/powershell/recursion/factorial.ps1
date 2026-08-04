# vybe-test: powershell/recursion/factorial
function Get-Factorial([int]$n) {
    if ($n -le 1) { return 1 }
    return $n * (Get-Factorial ($n - 1))
}
$result = Get-Factorial 7
if ($result -ne 5040) {
    Write-Host "FAIL: expected 5040, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
