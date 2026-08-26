# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_chained_to_sort_object
$items = @(
    [pscustomobject]@{ Val = 3 },
    [pscustomobject]@{ Val = 1 },
    [pscustomobject]@{ Val = 2 }
)
$sorted = @($items | Select-Object @{ N = "Cube"; E = { [math]::Pow($_.Val, 3) } } | Sort-Object -Property Cube)
if ($sorted[0].Cube -ne 1 -or $sorted[1].Cube -ne 8 -or $sorted[2].Cube -ne 27) {
    Write-Host "FAIL: Calculated property chained to Sort-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
