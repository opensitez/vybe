# vybe-test: powershell/pscustomobject_literals/pscustomobject_type_check
$obj = [pscustomobject]@{ A = 1 }
if (-not ($obj -is [PSCustomObject])) {
    Write-Host "FAIL: object is not [PSCustomObject]"
    exit 1
}
Write-Host "PASS"
exit 0
