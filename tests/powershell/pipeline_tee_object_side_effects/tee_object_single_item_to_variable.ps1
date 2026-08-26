# vybe-test: powershell/pipeline_tee_object_side_effects/tee_object_single_item_to_variable
$var = $null
$res = "Hello" | Tee-Object -Variable var
if ($res -ne "Hello" -or $var -ne "Hello") {
    Write-Host "FAIL: Tee-Object single item to variable failed"
    exit 1
}
Write-Host "PASS"
exit 0
