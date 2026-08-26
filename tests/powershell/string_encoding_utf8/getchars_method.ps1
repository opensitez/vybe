# vybe-test: powershell/string_encoding_utf8/getchars_method
$enc = [System.Text.Encoding]::UTF8
[byte[]]$bytes = @(88, 89, 90)
$chars = $enc.GetChars($bytes)
if ($chars.Length -ne 3 -or $chars[0] -ne [char]'X' -or $chars[2] -ne [char]'Z') {
    Write-Host "FAIL: GetChars method failed"
    exit 1
}
Write-Host "PASS"
exit 0
