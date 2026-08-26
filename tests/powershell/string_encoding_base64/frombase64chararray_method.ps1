# vybe-test: powershell/string_encoding_base64/frombase64chararray_method
$chars = "SGVsbG8=".ToCharArray()
$bytes = [System.Convert]::FromBase64CharArray($chars, 0, $chars.Length)
$str = [System.Text.Encoding]::UTF8.GetString($bytes)
if ($str -ne "Hello") {
    Write-Host "FAIL: FromBase64CharArray failed, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
