# vybe-test: powershell/type_unsigned_integers/uint64_in_math_functions
[uint64]$a = 1000000000
[uint64]$b = 2000000000
$max = [math]::Max($a, $b)
if ($max -ne $b) {
    Write-Host "FAIL: [math]::Max on uint64 failed"
    exit 1
}
Write-Host "PASS"
exit 0
