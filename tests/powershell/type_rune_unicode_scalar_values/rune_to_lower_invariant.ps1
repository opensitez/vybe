# vybe-test: powershell/type_rune_unicode_scalar_values/rune_to_lower_invariant
$rUpper = [System.Text.Rune]::new([char]'M')
$rLower = [System.Text.Rune]::ToLowerInvariant($rUpper)
if ($rLower.Value -ne [int][char]'m') { Write-Host "FAIL: Rune ToLowerInvariant failed"; exit 1 }
Write-Host "PASS"; exit 0
