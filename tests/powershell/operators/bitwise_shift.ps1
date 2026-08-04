# vybe-test: powershell/operators/bitwise_shift
$val = 1
$left4  = $val -shl 4   # 1 << 4 = 16
$right2 = 64 -shr 2     # 64 >> 2 = 16
if ($left4  -ne 16) { Write-Host "FAIL: shl: $left4";  exit 1 }
if ($right2 -ne 16) { Write-Host "FAIL: shr: $right2"; exit 1 }
# Combine: set bit 3, clear bit 1
$flags = 0b00001010     # bits 1 and 3
$flags = $flags -bor  (1 -shl 5)   # set bit 5 → 0b00101010 = 42
$flags = $flags -band (-bnot (1 -shl 1)) # clear bit 1 → 0b00101000 = 40
if ($flags -ne 40) { Write-Host "FAIL: flags $flags"; exit 1 }
Write-Host "PASS"
exit 0
