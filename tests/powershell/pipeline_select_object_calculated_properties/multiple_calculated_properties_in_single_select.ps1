# vybe-test: powershell/pipeline_select_object_calculated_properties/multiple_calculated_properties_in_single_select
$rect = [pscustomobject]@{ Width = 10; Height = 5 }
$res = $rect | Select-Object @{ N = "Area"; E = { $_.Width * $_.Height } }, @{ N = "Perimeter"; E = { 2 * ($_.Width + $_.Height) } }
if ($res.Area -ne 50 -or $res.Perimeter -ne 30) {
    Write-Host "FAIL: Multiple calculated properties in single select failed"
    exit 1
}
Write-Host "PASS"
exit 0
