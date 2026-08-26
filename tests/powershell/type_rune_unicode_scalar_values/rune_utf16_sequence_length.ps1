# vybe-test: powershell/type_rune_unicode_scalar_values/rune_utf16_sequence_length
$rBmp = [System.Text.Rune]::new([char]'A')
$rSurrogate = [System.Text.Rune]::new(0x1F600)
if ($rBmp.Utf16SequenceLength -ne 1 -or $rSurrogate.Utf16SequenceLength -ne 2) { Write-Host "FAIL: Rune Utf16SequenceLength failed"; exit 1 }
Write-Host "PASS"; exit 0
