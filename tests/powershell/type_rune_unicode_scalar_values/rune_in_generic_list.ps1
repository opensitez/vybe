# vybe-test: powershell/type_rune_unicode_scalar_values/rune_in_generic_list
$list = [System.Collections.Generic.List[System.Text.Rune]]::new()
$list.Add([System.Text.Rune]::new([char]'X'))
if ($list.Count -ne 1 -or $list[0].Value -ne [int][char]'X') { Write-Host "FAIL: Rune in List failed"; exit 1 }
Write-Host "PASS"; exit 0
