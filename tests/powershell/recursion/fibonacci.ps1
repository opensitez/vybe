# vybe-test: powershell/recursion/fibonacci
function Get-Fibonacci {
    param([int]$n)
    if ($n -le 1) {
        return $n
    }
    return (Get-Fibonacci ($n - 1)) + (Get-Fibonacci ($n - 2))
}

$result = Get-Fibonacci 7
if ($result -ne 13) {
    Write-Host "FAIL: expected 13, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
