# vybe-test: powershell/psvariable_objects/psvariable_in_function
function Get-LocalVarObj {
    $localVar = "ScopedData"
    return Get-Variable -Name "localVar"
}
$v = Get-LocalVarObj
if ($v.Value -ne "ScopedData") {
    Write-Host "FAIL: Get-Variable in function expected Value='ScopedData'"
    exit 1
}
Write-Host "PASS"
exit 0
