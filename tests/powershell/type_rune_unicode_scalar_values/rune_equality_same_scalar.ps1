# vybe-test: powershell/type_rune_unicode_scalar_values/rune_equality_same_scalar
$r1 = [System.Text.Rune]::new(0x1F600)
$r2 = [System.Text.Rune]::new(0x1F600)
if (-not $r1.Equals($r2)) { Write-Host "FAIL: Rune Equals failed"; exit 1 }
Write-Host "PASS"; exit 0
