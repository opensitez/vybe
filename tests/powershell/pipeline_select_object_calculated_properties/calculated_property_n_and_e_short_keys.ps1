# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_n_and_e_short_keys
$nums = @(1, 2, 3)
$res = @($nums | Select-Object @{ n = "Square"; e = { $_ * $_ } })
if ($res.Length -ne 3 -or $res[0].Square -ne 1 -or $res[2].Square -ne 9) {
    Write-Host "FAIL: Calculated property n/e short keys failed"
    exit 1
}
Write-Host "PASS"
exit 0
