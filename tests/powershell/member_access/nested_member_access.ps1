# vybe-test: powershell/member_access/nested_member_access
$obj = [pscustomobject]@{ Inner = [pscustomobject]@{ Value = 7 } }
if ($obj.Inner.Value -ne 7) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
