# vybe-test: powershell/pstypenames/pstypenames_custom_formatting_hook
$obj = [pscustomobject]@{ Value = 50 }
$obj.PSTypeNames.Insert(0, "CustomFormatType")
if ($obj.PSTypeNames[0] -ne "CustomFormatType") {
    Write-Host "FAIL: PSTypeNames property alias access expected CustomFormatType"
    exit 1
}
Write-Host "PASS"
exit 0
