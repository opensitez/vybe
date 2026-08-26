# vybe-test: powershell/string_encoding_hex_conversions/fromhexstring_invalid_character_exception
$caught = $false
try {
    $x = [System.Convert]::FromHexString("GG")
} catch [System.FormatException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected FormatException on non-hex characters"
    exit 1
}
Write-Host "PASS"
exit 0
