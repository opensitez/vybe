# vybe-test: powershell/type_rune_unicode_scalar_values/rune_creation_from_char
$r = [System.Text.Rune]::new([char]'A')
if ($r.Value -ne 65) { Write-Host "FAIL: Rune from char failed"; exit 1 }
Write-Host "PASS"; exit 0
