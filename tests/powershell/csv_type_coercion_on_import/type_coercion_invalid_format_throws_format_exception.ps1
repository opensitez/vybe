# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_invalid_format_throws_format_exception
$csv = "Num`nNotANumber"
$row = $csv | ConvertFrom-Csv
$caught = $false
try {
    $x = [int]::Parse($row.Num)
} catch [System.FormatException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected FormatException on invalid int coercion"
    exit 1
}
Write-Host "PASS"
exit 0
