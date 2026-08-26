# vybe-test: powershell/string_base85_and_base32_encoding/base32_hex_conversion_17
$bytes = [System.Text.Encoding]::UTF8.GetBytes("Test_17")
$hex = [System.Convert]::ToHexString($bytes)
$recovered = [System.Convert]::FromHexString($hex)
$str = [System.Text.Encoding]::UTF8.GetString($recovered)
if ($str -ne "Test_17") { Write-Host "FAIL: Hex string roundtrip failed"; exit 1 }
Write-Host "PASS"; exit 0
