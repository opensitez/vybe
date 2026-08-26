# vybe-test: powershell/string_encoding_hex_conversions/bitconverter_tostring_subarray_range
[byte[]]$bytes = @(0, 10, 20, 0)
$str = [System.BitConverter]::ToString($bytes, 1, 2)
if ($str -ne "0A-14") {
    Write-Host "FAIL: BitConverter ToString subarray failed, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
