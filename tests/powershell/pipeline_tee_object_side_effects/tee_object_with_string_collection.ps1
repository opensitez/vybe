# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_string_collection
$words = @("alpha", "beta", "gamma")
$t = $null
$res = @($words | Tee-Object -Variable t | ForEach-Object { $_.ToUpper() })
if ($res[0] -ne "ALPHA" -or $t[0] -ne "alpha") {
    Write-Host "FAIL: Tee-Object string collection failed"
    exit 1
}
Write-Host "PASS"
exit 0
