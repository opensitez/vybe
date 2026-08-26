# vybe-test: powershell/type_index_and_range_constructs/index_tostring_start_and_end
$iStart = [System.Index]::FromStart(4)
$iEnd = [System.Index]::FromEnd(2)
if ($iStart.ToString() -ne "4" -or $iEnd.ToString() -ne "^2") { Write-Host "FAIL: Index ToString failed"; exit 1 }
Write-Host "PASS"; exit 0
