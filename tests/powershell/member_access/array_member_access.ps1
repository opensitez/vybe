# vybe-test: powershell/member_access/array_member_access
$arr = 1,2,3
if ($arr.Length -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
