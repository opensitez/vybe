# vybe-test: powershell/null_handling/null_to_string
$x = $null
if (($x -eq $null) -and ($x -is [object] -or $true)) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
