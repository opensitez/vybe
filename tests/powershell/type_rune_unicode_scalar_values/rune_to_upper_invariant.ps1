# vybe-test: powershell/type_rune_unicode_scalar_values/rune_to_upper_invariant
$rLower = [System.Text.Rune]::new([char]'g')
$rUpper = [System.Text.Rune]::ToUpperInvariant($rLower)
if ($rUpper.Value -ne [int][char]'G') { Write-Host "FAIL: Rune ToUpperInvariant failed"; exit 1 }
Write-Host "PASS"; exit 0
