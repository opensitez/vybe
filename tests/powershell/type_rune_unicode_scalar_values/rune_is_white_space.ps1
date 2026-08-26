# vybe-test: powershell/type_rune_unicode_scalar_values/rune_is_white_space
$rSpace = [System.Text.Rune]::new([char]' ')
$rTab = [System.Text.Rune]::new([char]"`t")
$rChar = [System.Text.Rune]::new([char]'A')
if (-not [System.Text.Rune]::IsWhiteSpace($rSpace) -or -not [System.Text.Rune]::IsWhiteSpace($rTab) -or [System.Text.Rune]::IsWhiteSpace($rChar)) {
    Write-Host "FAIL: Rune IsWhiteSpace failed"; exit 1
}
Write-Host "PASS"; exit 0
