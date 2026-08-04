# vybe-test: powershell/assignment/hash_assignment
$hash = @{ a = 1 }
if ($hash.a -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
