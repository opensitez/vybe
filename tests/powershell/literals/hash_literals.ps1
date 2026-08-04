# vybe-test: powershell/literals/hash_literals
$hash = @{ a = 1; b = 2 }
if ($hash['a'] -ne 1 -or $hash.b -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
