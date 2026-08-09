# vybe-test: powershell/pscustomobject_literals/pscustomobject_property_types
$obj = [pscustomobject]@{
    IntVal = 5
    StrVal = "text"
    ArrVal = @(1, 2)
}
if (-not ($obj.IntVal -is [int])) {
    Write-Host "FAIL: IntVal is not [int]"
    exit 1
}
if (-not ($obj.ArrVal -is [array])) {
    Write-Host "FAIL: ArrVal is not [array]"
    exit 1
}
Write-Host "PASS"
exit 0
