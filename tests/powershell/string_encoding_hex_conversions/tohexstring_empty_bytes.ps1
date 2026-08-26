# vybe-test: powershell/string_encoding_hex_conversions/tohexstring_empty_bytes
[byte[]]$empty = @()
$hex = [System.Convert]::ToHexString($empty)
if ($hex -ne "") {
    Write-Host "FAIL: ToHexString empty bytes failed"
    exit 1
}
Write-Host "PASS"
exit 0
