# vybe-test: powershell/type_conversion/convert_base_hex
$hex = [Convert]::ToInt32("FF", 16)
if ($hex -ne 255) { Write-Host "FAIL: hex FF should be 255, got $hex"; exit 1 }
$back = [Convert]::ToString(255, 16)
if ($back -ne "ff") { Write-Host "FAIL: 255 to hex should be 'ff', got $back"; exit 1 }
Write-Host "PASS"
exit 0
