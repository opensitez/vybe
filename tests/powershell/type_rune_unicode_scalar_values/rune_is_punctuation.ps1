# vybe-test: powershell/type_rune_unicode_scalar_values/rune_is_punctuation
$rDot = [System.Text.Rune]::new([char]'.')
$rLetter = [System.Text.Rune]::new([char]'A')
if (-not [System.Text.Rune]::IsPunctuation($rDot) -or [System.Text.Rune]::IsPunctuation($rLetter)) {
    Write-Host "FAIL: Rune IsPunctuation failed"; exit 1
}
Write-Host "PASS"; exit 0
