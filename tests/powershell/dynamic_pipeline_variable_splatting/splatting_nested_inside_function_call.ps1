# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_nested_inside_function_call
function Inner-Target {
    param([int]$X, [int]$Y)
    return $X + $Y
}
function Outer-Wrapper {
    param([hashtable]$InnerParams)
    return Inner-Target @InnerParams
}
$res = Outer-Wrapper -InnerParams @{ X = 15; Y = 25 }
if ($res -ne 40) {
    Write-Host "FAIL: Nested splatting inside function call failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
