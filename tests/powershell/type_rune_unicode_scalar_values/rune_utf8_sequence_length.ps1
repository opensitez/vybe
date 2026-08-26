# vybe-test: powershell/type_rune_unicode_scalar_values/rune_utf8_sequence_length
$r1 = [System.Text.Rune]::new([char]'A')
$r2 = [System.Text.Rune]::new(0x00E9)
$r3 = [System.Text.Rune]::new(0x4E2D)
$r4 = [System.Text.Rune]::new(0x1F600)
if ($r1.Utf8SequenceLength -ne 1 -or $r2.Utf8SequenceLength -ne 2 -or $r3.Utf8SequenceLength -ne 3 -or $r4.Utf8SequenceLength -ne 4) {
    Write-Host "FAIL: Rune Utf8SequenceLength failed"
    exit 1
}
Write-Host "PASS"
exit 0
