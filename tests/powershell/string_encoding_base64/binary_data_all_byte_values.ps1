# vybe-test: powershell/string_encoding_base64/binary_data_all_byte_values
[byte[]]$allBytes = New-Object byte[] 256
for ($i = 0; $i -lt 256; $i++) { $allBytes[$i] = [byte]$i }
$b64 = [System.Convert]::ToBase64String($allBytes)
$decoded = [System.Convert]::FromBase64String($b64)
$mismatch = $false
for ($i = 0; $i -lt 256; $i++) {
    if ($decoded[$i] -ne $allBytes[$i]) { $mismatch = $true; break }
}
if ($mismatch -or $decoded.Length -ne 256) {
    Write-Host "FAIL: All 256 byte values base64 roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
