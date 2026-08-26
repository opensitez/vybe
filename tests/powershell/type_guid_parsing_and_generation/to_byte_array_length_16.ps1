# vybe-test: powershell/type_guid_parsing_and_generation/to_byte_array_length_16
$g = [guid]::NewGuid()
$bytes = $g.ToByteArray()
if ($bytes.Length -ne 16) {
    Write-Host "FAIL: Guid byte array must have length 16"
    exit 1
}
Write-Host "PASS"
exit 0
