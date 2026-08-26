# vybe-test: powershell/type_rune_unicode_scalar_values/rune_trycreate_valid_and_invalid
$outRune = [System.Text.Rune]::new([char]'X')
$ok1 = [System.Text.Rune]::TryCreate(0x1F600, [ref]$outRune)
$ok2 = [System.Text.Rune]::TryCreate(0xD800, [ref]$outRune) # surrogate is invalid scalar
if (-not $ok1 -or $ok2) { Write-Host "FAIL: Rune TryCreate failed"; exit 1 }
Write-Host "PASS"; exit 0
