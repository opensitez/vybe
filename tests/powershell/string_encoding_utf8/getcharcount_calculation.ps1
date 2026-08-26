# vybe-test: powershell/string_encoding_utf8/getcharcount_calculation
$enc = [System.Text.Encoding]::UTF8
[byte[]]$bytes = @(65, 66, 0xC3, 0xA9) # A, B, é
$count = $enc.GetCharCount($bytes)
if ($count -ne 3) {
    Write-Host "FAIL: GetCharCount expected 3, got $count"
    exit 1
}
Write-Host "PASS"
exit 0
