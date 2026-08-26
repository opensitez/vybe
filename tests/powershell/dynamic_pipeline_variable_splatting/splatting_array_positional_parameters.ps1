# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_array_positional_parameters
function Target-PosFunc {
    param([int]$A, [int]$B, [int]$C)
    return ($A + $B + $C)
}
$arr = @(10, 20, 30)
$res = Target-PosFunc @arr
if ($res -ne 60) {
    Write-Host "FAIL: Array positional splatting failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
