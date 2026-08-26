# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_returning_array
$item = [pscustomobject]@{ Csv = "a,b,c" }
$res = $item | Select-Object @{ N = "Parts"; E = { $_.Csv.Split(',') } }
if ($res.Parts.Length -ne 3 -or $res.Parts[0] -ne "a" -or $res.Parts[2] -ne "c") {
    Write-Host "FAIL: Calculated property returning array failed"
    exit 1
}
Write-Host "PASS"
exit 0
