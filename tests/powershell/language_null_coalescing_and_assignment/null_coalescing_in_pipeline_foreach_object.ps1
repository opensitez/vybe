# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_in_pipeline_foreach_object
$items = @("First", $null, "Third")
$coalesced = @($items | ForEach-Object { $_ ?? "Placeholder" })
if ($coalesced[0] -ne "First" -or $coalesced[1] -ne "Placeholder" -or $coalesced[2] -ne "Third") {
    Write-Host "FAIL: ?? in pipeline ForEach-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
