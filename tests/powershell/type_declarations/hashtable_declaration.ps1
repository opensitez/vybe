# vybe-test: powershell/type_declarations/hashtable_declaration
[hashtable]$x = @{ a = 1 }
if ($x.a -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
