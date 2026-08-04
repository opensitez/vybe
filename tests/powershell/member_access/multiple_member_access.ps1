# vybe-test: powershell/member_access/multiple_member_access
$obj = [pscustomobject]@{ Inner = [pscustomobject]@{ Value = 'ok' } }
if ($obj.Inner.Value -ne 'ok') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
