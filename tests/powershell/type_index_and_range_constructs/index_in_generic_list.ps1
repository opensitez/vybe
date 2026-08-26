# vybe-test: powershell/type_index_and_range_constructs/index_in_generic_list
$list = [System.Collections.Generic.List[System.Index]]::new()
$list.Add([System.Index]::FromStart(10))
if ($list.Count -ne 1 -or $list[0].Value -ne 10) { Write-Host "FAIL: Index in List failed"; exit 1 }
Write-Host "PASS"; exit 0
