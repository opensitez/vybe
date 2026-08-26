# vybe-test: powershell/type_rune_unicode_scalar_values/rune_is_bmp_property
$rBmp = [System.Text.Rune]::new([char]'X')
$rNonBmp = [System.Text.Rune]::new(0x1F600)
if (-not $rBmp.IsBmp -or $rNonBmp.IsBmp) { Write-Host "FAIL: Rune IsBmp failed"; exit 1 }
Write-Host "PASS"; exit 0
