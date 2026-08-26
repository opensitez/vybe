# vybe-test: powershell/csv_header_manipulation/duplicate_headers_in_csv_handled_or_renamed
$csv = @"
Col,Col
1,2
"@
$caughtOrRenamed = $false
try {
    $rows = @($csv | ConvertFrom-Csv)
    $caughtOrRenamed = ($rows.Length -eq 1)
} catch {
    $caughtOrRenamed = $true # Duplicate header throwing is also acceptable
}
if (-not $caughtOrRenamed) {
    Write-Host "FAIL: Duplicate header in CSV failed"
    exit 1
}
Write-Host "PASS"
exit 0
