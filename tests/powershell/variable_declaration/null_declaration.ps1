# vybe-test: powershell/variable_declaration/null_declaration
$x = $null
if ($x -ne $null) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
