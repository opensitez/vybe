# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_empty_hashtable_is_noop
function Target-NoopSplat {
    param([string]$DefaultVal = "Default")
    return $DefaultVal
}
$empty = @{}
$res = Target-NoopSplat @empty
if ($res -ne "Default") {
    Write-Host "FAIL: Splatting empty hashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
