# vybe-test: powershell/string_encoding_hex_conversions/bitconverter_tostring_with_hyphens
[byte[]]$bytes = @(1, 2, 15, 16)
$str = [System.BitConverter]::ToString($bytes)
if ($str -ne "01-02-0F-10") {
    Write-Host "FAIL: BitConverter ToString failed, expected '01-02-0F-10', got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
