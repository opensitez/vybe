# vybe-test: powershell/recursion/fibonacci
function Get-Fib([int]$n) {
    if ($n -le 0) { return 0 }
    if ($n -eq 1) { return 1 }
    return (Get-Fib ($n - 1)) + (Get-Fib ($n - 2))
}
$expected = @(0,1,1,2,3,5,8,13,21,34)
for ($i = 0; $i -lt 10; $i++) {
    $got = Get-Fib $i
    if ($got -ne $expected[$i]) {
        Write-Host "FAIL: fib($i) expected $($expected[$i]) got $got"
        exit 1
    }
}
Write-Host "PASS"
exit 0
