# vybe-test: powershell/string_encoding_base64/tobase64string_with_line_breaks_option
[byte[]]$large = New-Object byte[] 100
for ($i=0; $i -lt 100; $i++) { $large[$i] = [byte]($i % 256) }
$b64 = [System.Convert]::ToBase64String($large, [System.Base64FormattingOptions]::InsertLineBreaks)
if (-not ($b64.Contains("`r`n") -or $b64.Contains("`n"))) {
    Write-Host "FAIL: Base64 with InsertLineBreaks failed"
    exit 1
}
Write-Host "PASS"
exit 0
