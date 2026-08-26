# vybe-test: powershell/type_unsigned_integers/byte_addition_and_coercion
[byte]$b1 = 200
[byte]$b2 = 50
$sum = $b1 + $b2
if ($sum -ne 250) {
    Write-Host "FAIL: byte sum mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
