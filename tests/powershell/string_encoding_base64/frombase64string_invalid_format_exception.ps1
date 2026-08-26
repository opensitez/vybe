# vybe-test: powershell/string_encoding_base64/frombase64string_invalid_format_exception
$caught = $false
try {
    $x = [System.Convert]::FromBase64String("Invalid!!!Base64===")
} catch [System.FormatException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected FormatException on invalid base64"
    exit 1
}
Write-Host "PASS"
exit 0
