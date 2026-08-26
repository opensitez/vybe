# vybe-test: powershell/type_rune_unicode_scalar_values/rune_get_numeric_value
$r = [System.Text.Rune]::new([char]'5')
$val = [System.Text.Rune]::GetNumericValue($r)
if ($val -ne 5.0) { Write-Host "FAIL: Rune GetNumericValue failed"; exit 1 }
Write-Host "PASS"; exit 0
