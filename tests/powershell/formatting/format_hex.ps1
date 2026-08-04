# vybe-test: powershell/formatting/format_hex_binary
$hex = "{0:X}" -f 255
if ($hex -ne "FF") { Write-Host "FAIL: hex $hex"; exit 1 }
$hex4 = "{0:X4}" -f 10
if ($hex4 -ne "000A") { Write-Host "FAIL: hex4 $hex4"; exit 1 }
Write-Host "PASS"
exit 0
