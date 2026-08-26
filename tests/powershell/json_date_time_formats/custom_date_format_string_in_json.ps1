# vybe-test: powershell/json_date_time_formats/custom_date_format_string_in_json
$dt = [datetime]::ParseExact("2026-08-26 15:45:00", "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
$customStr = $dt.ToString("dd/MM/yyyy HH:mm")
$json = @{ Formatted = $customStr } | ConvertTo-Json
$obj = $json | ConvertFrom-Json
if ($obj.Formatted -ne "26/08/2026 15:45") {
    Write-Host "FAIL: Custom date format string in JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
