# vybe-test: powershell/type_rune_unicode_scalar_values/rune_is_control_character
$rNull = [System.Text.Rune]::new([char]"`0")
$rChar = [System.Text.Rune]::new([char]'A')
if (-not [System.Text.Rune]::IsControl($rNull) -or [System.Text.Rune]::IsControl($rChar)) { Write-Host "FAIL: Rune IsControl failed"; exit 1 }
Write-Host "PASS"; exit 0
