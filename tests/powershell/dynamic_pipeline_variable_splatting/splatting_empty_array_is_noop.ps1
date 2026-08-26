# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_empty_array_is_noop
function Target-EmptyArrSplat {
    param([string]$DefaultVal = "Default")
    return $DefaultVal
}
$empty = @()
$res = Target-EmptyArrSplat @empty
if ($res -ne "Default") {
    Write-Host "FAIL: Splatting empty array failed"
    exit 1
}
Write-Host "PASS"
exit 0
