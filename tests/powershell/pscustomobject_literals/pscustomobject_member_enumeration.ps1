# vybe-test: powershell/pscustomobject_literals/pscustomobject_member_enumeration
$obj = [pscustomobject]@{ P1 = "A"; P2 = "B" }
$names = $obj.psobject.Properties | ForEach-Object { $_.Name }
if ($names[0] -ne "P1" -or $names[1] -ne "P2") {
    Write-Host "FAIL: property enumeration expected P1, P2, got $($names -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
