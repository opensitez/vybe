# vybe-test: powershell/member_access/members_on_null
$obj = $null
if ($obj?.Value -ne $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
