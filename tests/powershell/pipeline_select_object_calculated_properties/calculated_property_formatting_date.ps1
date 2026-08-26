# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_formatting_date
$d = [datetime]::Parse("2026-08-26")
$res = $d | Select-Object @{ N = "IsoDate"; E = { $_.ToString("yyyy-MM-dd") } }
if ($res.IsoDate -ne "2026-08-26") {
    Write-Host "FAIL: Calculated property date formatting failed"
    exit 1
}
Write-Host "PASS"
exit 0
