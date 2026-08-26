# vybe-test: powershell/string_encoding_base64/tryfrombase64chars_method
$chars = "VGVzdA==".ToCharArray() # "Test"
[byte[]]$bytes = New-Object byte[] 4
$written = 0
$ok = [System.Convert]::TryFromBase64Chars($chars, $bytes, [ref]$written)
$str = [System.Text.Encoding]::UTF8.GetString($bytes)
if (-not $ok -or $written -ne 4 -or $str -ne "Test") {
    Write-Host "FAIL: TryFromBase64Chars failed"
    exit 1
}
Write-Host "PASS"
exit 0
