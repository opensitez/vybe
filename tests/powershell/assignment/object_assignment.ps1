# vybe-test: powershell/assignment/object_assignment
$obj = [pscustomobject]@{ Value = 1 }
if ($obj.Value -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
