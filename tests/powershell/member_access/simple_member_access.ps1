# vybe-test: powershell/member_access/simple_member_access
$obj = [pscustomobject]@{ Value = 5 }
if ($obj.Value -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
