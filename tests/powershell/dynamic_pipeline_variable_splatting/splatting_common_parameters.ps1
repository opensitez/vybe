# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_common_parameters
function Target-CommonParamCheck {
    [CmdletBinding()]
    param([string]$Msg)
    Write-Verbose $Msg -Verbose
    return "Done"
}
$p = @{ Msg = "TestMsg"; Verbose = $true }
$res = Target-CommonParamCheck @p
if ($res -ne "Done") {
    Write-Host "FAIL: Splatting common parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
