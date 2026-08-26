# vybe-test: powershell/string_encoding_utf8/empty_bytes_getstring
[byte[]]$empty = @()
$str = [System.Text.Encoding]::UTF8.GetString($empty)
if ($str -ne "") {
    Write-Host "FAIL: Empty bytes GetString must return empty string"
    exit 1
}
Write-Host "PASS"
exit 0
