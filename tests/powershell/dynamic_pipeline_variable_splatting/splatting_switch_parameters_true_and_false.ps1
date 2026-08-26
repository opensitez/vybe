# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_switch_parameters_true_and_false
function Target-SwitchFunc {
    param([switch]$Force, [switch]$Verbose)
    return "F:$($Force.IsPresent),V:$($Verbose.IsPresent)"
}
$p1 = @{ Force = $true; Verbose = $false }
$res1 = Target-SwitchFunc @p1
if ($res1 -ne "F:True,V:False") {
    Write-Host "FAIL: Splatting switch parameters failed, got '$res1'"
    exit 1
}
Write-Host "PASS"
exit 0
