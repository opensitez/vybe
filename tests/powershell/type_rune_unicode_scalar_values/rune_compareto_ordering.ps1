# vybe-test: powershell/type_rune_unicode_scalar_values/rune_compareto_ordering
$r1 = [System.Text.Rune]::new([char]'A')
$r2 = [System.Text.Rune]::new([char]'B')
if ($r1.CompareTo($r2) -ge 0 -or $r2.CompareTo($r1) -le 0) { Write-Host "FAIL: Rune CompareTo failed"; exit 1 }
Write-Host "PASS"; exit 0
