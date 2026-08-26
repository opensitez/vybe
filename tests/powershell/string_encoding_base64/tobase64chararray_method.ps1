# vybe-test: powershell/string_encoding_base64/tobase64chararray_method
$bytes = [System.Text.Encoding]::UTF8.GetBytes("Hi")
[char[]]$outChars = New-Object char[] 4
$written = [System.Convert]::ToBase64CharArray($bytes, 0, 2, $outChars, 0)
$str = -join $outChars
if ($written -ne 4 -or $str -ne "SGk=") {
    Write-Host "FAIL: ToBase64CharArray failed, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
