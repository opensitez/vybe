# vybe-test: powershell/ref_parameters/ref_param_hashtable_mutation
function Update-Hash([ref]$hRef) {
    $hRef.Value["NewKey"] = "NewValue"
}
$map = @{ Existing = "Old" }
Update-Hash ([ref]$map)
if ($map["NewKey"] -ne "NewValue") {
    Write-Host "FAIL: hashtable [ref] mutation expected NewValue"
    exit 1
}
Write-Host "PASS"
exit 0
