# vybe-test: powershell/psvariable_objects/psvariable_subexpression
$myVar = "SubVal"
$varObj = Get-Variable -Name "myVar"
$msg = "Var: $( $varObj.Value )"
if ($msg -ne "Var: SubVal") {
    Write-Host "FAIL: PSVariable in subexpression expected 'Var: SubVal', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
