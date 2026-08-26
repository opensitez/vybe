# vybe-test: powershell/pipeline_chain_operators_and_or/chained_and_operators_three_commands
$log = [System.Collections.Generic.List[string]]::new()
function Step1 { $log.Add("1"); $true }
function Step2 { $log.Add("2"); $true }
function Step3 { $log.Add("3"); $true }
Step1 && Step2 && Step3
if ($log.Count -ne 3 -or $log[2] -ne "3") {
    Write-Host "FAIL: Chained && across 3 commands failed"
    exit 1
}
Write-Host "PASS"
exit 0
