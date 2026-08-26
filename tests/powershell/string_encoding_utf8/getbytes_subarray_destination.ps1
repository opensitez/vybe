# vybe-test: powershell/string_encoding_utf8/getbytes_subarray_destination
$enc = [System.Text.Encoding]::UTF8
[byte[]]$dest = New-Object byte[] 5
$written = $enc.GetBytes("Hi", 0, 2, $dest, 1) # write at index 1
if ($written -ne 2 -or $dest[1] -ne 72 -or $dest[2] -ne 105) {
    Write-Host "FAIL: GetBytes to destination array failed"
    exit 1
}
Write-Host "PASS"
exit 0
