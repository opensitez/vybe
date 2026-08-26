# vybe-test: powershell/string_base85_and_base32_encoding/base32_hex_conversion_19
$bytes = [System.Text.Encoding]::UTF8.GetBytes("Test_19")
$hex = [System.Convert]::ToHexString($bytes)
$recovered = [System.Convert]::FromHexString($hex)
$str = [System.Text.Encoding]::UTF8.GetString($recovered)
if ($str -ne "Test_19") { Write-Host "FAIL: Hex string roundtrip failed"; exit 1 }
Write-Host "PASS"; exit 0
