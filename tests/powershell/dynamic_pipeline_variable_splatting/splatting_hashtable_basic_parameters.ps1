# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_hashtable_basic_parameters
function Target-Func {
    param([string]$First, [string]$Last)
    return "$First $Last"
}
$params = @{ First = "Alice"; Last = "Smith" }
$res = Target-Func @params
if ($res -ne "Alice Smith") {
    Write-Host "FAIL: Basic hashtable splatting failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
