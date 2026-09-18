# vybe-test: powershell/recursion/factorial
function Get-Factorial {
    param([int]$n)
    if ($n -le 1) {
        return 1
    }
    return $n * (Get-Factorial ($n - 1))
}

$result = Get-Factorial 5
if ($result -ne 120) {
    Write-Host "FAIL: expected 120, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
