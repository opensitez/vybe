# vybe-test: powershell/type_rune_unicode_scalar_values/rune_hashcode_consistency
$r1 = [System.Text.Rune]::new(0x00E9)
$r2 = [System.Text.Rune]::new(0x00E9)
if ($r1.GetHashCode() -ne $r2.GetHashCode()) { Write-Host "FAIL: Rune HashCode failed"; exit 1 }
Write-Host "PASS"; exit 0
