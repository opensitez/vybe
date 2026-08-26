# vybe-test: powershell/type_unsigned_integers/parse_uint32_from_hex_string
$val = [uint32]::Parse("DEADBEEF", [System.Globalization.NumberStyles]::HexNumber)
if ($val -ne 3735928559) {
    Write-Host "FAIL: DEADBEEF parse mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
