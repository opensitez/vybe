# vybe-test: powershell/pscustomobject_literals/pscustomobject_empty
$obj = [pscustomobject]@{}
$count = $obj.psobject.Properties.Name.Count
if ($count -ne 0) {
    Write-Host "FAIL: empty PSCustomObject property count expected 0, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
