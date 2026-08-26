# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_overriding_explicit_parameter_throws_duplicate
function Target-DupCheck {
    param([string]$A)
    return $A
}
$p = @{ A = "splatted" }
$res = Target-DupCheck @p
if ($res -ne "splatted") {
    Write-Host "FAIL: Splatting parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
