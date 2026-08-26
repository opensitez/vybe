# vybe-test: powershell/string_encoding_hex_conversions/bitconverter_toboolean_from_byte
[byte[]]$trueByte = @(1)
[byte[]]$falseByte = @(0)
$t = [System.BitConverter]::ToBoolean($trueByte, 0)
$f = [System.BitConverter]::ToBoolean($falseByte, 0)
if ($t -ne $true -or $f -ne $false) {
    Write-Host "FAIL: BitConverter ToBoolean failed"
    exit 1
}
Write-Host "PASS"
exit 0
