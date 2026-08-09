# vybe-test: powershell/psvariable_objects/psvariable_get_variable_cmdlet
$TargetVar = "TargetData"
$varObj = Get-Variable -Name "TargetVar"
if ($varObj.Value -ne "TargetData") {
    Write-Host "FAIL: Get-Variable -Name TargetVar expected Value='TargetData'"
    exit 1
}
Write-Host "PASS"
exit 0
