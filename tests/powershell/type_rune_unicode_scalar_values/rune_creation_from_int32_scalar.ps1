# vybe-test: powershell/type_rune_unicode_scalar_values/rune_creation_from_int32_scalar
$r = [System.Text.Rune]::new(0x1F600) # 😀 grinning face
if ($r.Value -ne 0x1F600 -or -not $r.IsBmp -eq $false) { Write-Host "FAIL: Rune from int32 failed"; exit 1 }
Write-Host "PASS"; exit 0
