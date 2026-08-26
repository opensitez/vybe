# vybe-test: powershell/type_unsigned_integers/to_hex_string_formatting
[byte]$b = 255
if ($b.ToString("X2") -ne "FF") {
    Write-Host "FAIL: byte hex format failed"
    exit 1
}
Write-Host "PASS"
exit 0
