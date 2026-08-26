# vybe-test: powershell/string_encoding_base64/base64_with_spaces_ignored_in_standard_decode
$clean = "SGVsbG8="
$spaced = " SGVs bG8= "
$b1 = [System.Convert]::FromBase64String($clean)
$b2 = [System.Convert]::FromBase64String($spaced)
if ($b1.Length -ne $b2.Length -or $b1[0] -ne $b2[0]) {
    Write-Host "FAIL: Base64 with whitespace decode failed"
    exit 1
}
Write-Host "PASS"
exit 0
