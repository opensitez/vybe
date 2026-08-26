# vybe-test: powershell/type_guid_parsing_and_generation/from_byte_array_reconstruction
$orig = [guid]::NewGuid()
$bytes = $orig.ToByteArray()
$reconstructed = [guid]::new($bytes)
if ($orig -ne $reconstructed) {
    Write-Host "FAIL: Reconstructed GUID does not match original"
    exit 1
}
Write-Host "PASS"
exit 0
