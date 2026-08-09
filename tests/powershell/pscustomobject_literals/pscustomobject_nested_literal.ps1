# vybe-test: powershell/pscustomobject_literals/pscustomobject_nested_literal
$obj = [pscustomobject]@{
    Outer = [pscustomobject]@{ Inner = "Deep" }
}
if ($obj.Outer.Inner -ne "Deep") {
    Write-Host "FAIL: Outer.Inner expected Deep, got $($obj.Outer.Inner)"
    exit 1
}
Write-Host "PASS"
exit 0
