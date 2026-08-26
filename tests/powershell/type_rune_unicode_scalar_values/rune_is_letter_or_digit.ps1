# vybe-test: powershell/type_rune_unicode_scalar_values/rune_is_letter_or_digit
$rL = [System.Text.Rune]::new([char]'K')
$rD = [System.Text.Rune]::new([char]'7')
$rP = [System.Text.Rune]::new([char]'!')
if (-not [System.Text.Rune]::IsLetterOrDigit($rL) -or -not [System.Text.Rune]::IsLetterOrDigit($rD) -or [System.Text.Rune]::IsLetterOrDigit($rP)) {
    Write-Host "FAIL: Rune IsLetterOrDigit failed"; exit 1
}
Write-Host "PASS"; exit 0
