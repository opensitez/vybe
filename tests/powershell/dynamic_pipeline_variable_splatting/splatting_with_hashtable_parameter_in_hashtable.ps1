# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_with_hashtable_parameter_in_hashtable
function Target-HtParam {
    param([hashtable]$Config)
    return $Config["key"]
}
$p = @{ Config = @{ key = "innerVal" } }
$res = Target-HtParam @p
if ($res -ne "innerVal") {
    Write-Host "FAIL: Splatting hashtable parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
