# vybe-test: powershell/type_rune_unicode_scalar_values/rune_replacement_char_constant
$rep = [System.Text.Rune]::ReplacementChar
if ($rep.Value -ne 0xFFFD) { Write-Host "FAIL: Rune ReplacementChar failed"; exit 1 }
Write-Host "PASS"; exit 0
