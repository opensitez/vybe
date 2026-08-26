# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_case_insensitive_parameter_names
function Target-CaseSplat {
    param([string]$UserName)
    return $UserName
}
$p = @{ username = "alice" }
$res = Target-CaseSplat @p
if ($res -ne "alice") {
    Write-Host "FAIL: Case-insensitive parameter names in splatted hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
