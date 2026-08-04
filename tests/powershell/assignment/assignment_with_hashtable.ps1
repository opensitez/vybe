# vybe-test: powershell/assignment/assignment_with_hashtable
$hash = @{ a = 1 }
$hash['b'] = 2
if ($hash.b -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
