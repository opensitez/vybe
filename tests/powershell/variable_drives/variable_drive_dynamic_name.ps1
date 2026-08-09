# vybe-test: powershell/variable_drives/variable_drive_dynamic_name
$varName = "dynVar"
Set-Variable -Name $varName -Value "DynamicData"
$res = Get-Variable -Name $varName -ValueOnly
if ($res -ne "DynamicData") {
    Write-Host "FAIL: Get-Variable dynamic lookup expected 'DynamicData', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
