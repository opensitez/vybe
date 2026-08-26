# vybe-test: powershell/string_data_parser/invalid_line_without_equals_throws_or_errors
$caught = $false
try {
    $x = ConvertFrom-StringData -StringData "invalid_line_without_equals" -ErrorAction Stop
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected error on line missing equals sign"
    exit 1
}
Write-Host "PASS"
exit 0
