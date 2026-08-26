# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_with_group_object_downstream
$sideWords = $null
$groups = @("cat", "car", "dog", "cow" | Tee-Object -Variable sideWords | Group-Object { $_.Substring(0,1) })
if ($groups.Count -ne 2 -or $sideWords.Count -ne 4) {
    Write-Host "FAIL: Tee-Object with Group-Object downstream failed"
    exit 1
}
Write-Host "PASS"
exit 0
