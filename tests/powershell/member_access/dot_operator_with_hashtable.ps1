# vybe-test: powershell/member_access/dot_operator_with_hashtable
$hash = @{ a = 1 }
if ($hash.a -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
