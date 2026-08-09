# vybe-test: powershell/pipeline_chaining/chain_command_success_exit
$out = @()
(Write-Output "Step1") && ($script:out += "Step2")
if ($out.Count -ne 1 -or $out[0] -ne "Step2") {
    Write-Host "FAIL: pipeline command success && execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
