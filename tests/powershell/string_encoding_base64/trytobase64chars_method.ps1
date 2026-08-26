# vybe-test: powershell/string_encoding_base64/trytobase64chars_method
$bytes = [System.Text.Encoding]::UTF8.GetBytes("ABC")
[char[]]$chars = New-Object char[] 4
$written = 0
$ok = [System.Convert]::TryToBase64Chars($bytes, $chars, [ref]$written)
$str = -join $chars
if (-not $ok -or $written -ne 4 -or $str -ne "QUJD") {
    Write-Host "FAIL: TryToBase64Chars failed, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
