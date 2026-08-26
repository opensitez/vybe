# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_parameter_aliases_in_hashtable
function Target-AliasSplat {
    param([Alias("CN")][string]$ComputerName)
    return $ComputerName
}
$p = @{ CN = "server01" }
$res = Target-AliasSplat @p
if ($res -ne "server01") {
    Write-Host "FAIL: Parameter alias in splatted hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
