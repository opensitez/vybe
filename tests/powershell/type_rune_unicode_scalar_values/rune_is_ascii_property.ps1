# vybe-test: powershell/type_rune_unicode_scalar_values/rune_is_ascii_property
$rAscii = [System.Text.Rune]::new([char]'Z')
$rNonAscii = [System.Text.Rune]::new(0x00E9) # é
if (-not $rAscii.IsAscii -or $rNonAscii.IsAscii) { Write-Host "FAIL: Rune IsAscii failed"; exit 1 }
Write-Host "PASS"; exit 0
