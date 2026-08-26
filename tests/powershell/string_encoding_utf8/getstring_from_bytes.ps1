# vybe-test: powershell/string_encoding_utf8/getstring_from_bytes
[byte[]]$bytes = @(72, 101, 108, 108, 111)
$str = [System.Text.Encoding]::UTF8.GetString($bytes)
if ($str -ne "Hello") {
    Write-Host "FAIL: UTF8 GetString failed, got $str"
    exit 1
}
Write-Host "PASS"
exit 0
