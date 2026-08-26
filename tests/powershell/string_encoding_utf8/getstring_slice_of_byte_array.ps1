# vybe-test: powershell/string_encoding_utf8/getstring_slice_of_byte_array
[byte[]]$bytes = @(0, 72, 105, 0) # null, H, i, null
$str = [System.Text.Encoding]::UTF8.GetString($bytes, 1, 2)
if ($str -ne "Hi") {
    Write-Host "FAIL: GetString slice failed, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
