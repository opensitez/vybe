# vybe-test: powershell/type_rune_unicode_scalar_values/rune_tostring_representation
$r = [System.Text.Rune]::new(0x1F600)
$str = $r.ToString()
if ($str.Length -ne 2) { Write-Host "FAIL: Rune ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
