# vybe-test: powershell/psvariable_objects/psvariable_clear_variable_cmdlet
$ToClear = "DataToBeCleared"
Clear-Variable -Name "ToClear"
if ($ToClear -ne $null) {
    Write-Host "FAIL: Clear-Variable expected \$ToClear=null, got $ToClear"
    exit 1
}
Write-Host "PASS"
exit 0
