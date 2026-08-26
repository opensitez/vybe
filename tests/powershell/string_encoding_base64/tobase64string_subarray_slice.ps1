# vybe-test: powershell/string_encoding_base64/tobase64string_subarray_slice
[byte[]]$all = @(0, 65, 66, 67, 0) # null, A, B, C, null
$b64 = [System.Convert]::ToBase64String($all, 1, 3)
if ($b64 -ne "QUJD") { # ABC -> QUJD
    Write-Host "FAIL: Base64 subarray slice failed, got '$b64'"
    exit 1
}
Write-Host "PASS"
exit 0
