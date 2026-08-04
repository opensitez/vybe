# vybe-test: powershell/member_access/pscustomobject_property
$obj = [pscustomobject]@{ Name = 'x' }
if ($obj.Name -ne 'x') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
