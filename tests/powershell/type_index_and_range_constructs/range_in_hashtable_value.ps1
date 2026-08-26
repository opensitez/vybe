# vybe-test: powershell/type_index_and_range_constructs/range_in_hashtable_value
$h = @{ Selection = [System.Range]::All }
if ($h.Selection -ne [System.Range]::All) { Write-Host "FAIL: Range in hashtable failed"; exit 1 }
Write-Host "PASS"; exit 0
