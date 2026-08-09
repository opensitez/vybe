# vybe-test: powershell/variable_drives/variable_drive_provider_item
$itemVar = "ProviderVal"
$val = (Get-Item "variable:itemVar").Value
if ($val -ne "ProviderVal") {
    Write-Host "FAIL: Get-Item variable:itemVar expected 'ProviderVal', got '$val'"
    exit 1
}
Write-Host "PASS"
exit 0
