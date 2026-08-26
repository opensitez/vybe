# vybe-test: powershell/string_encoding_hex_conversions/fromhexstring_odd_length_exception
$caught = $false
try {
    $x = [System.Convert]::FromHexString("ABC")
} catch [System.FormatException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected FormatException on odd-length hex string"
    exit 1
}
Write-Host "PASS"
exit 0
