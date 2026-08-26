# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_with_pipeline_on_left
$log = [System.Collections.Generic.List[string]]::new()
function NextStep { $log.Add("Next") }
@(1, 2, 3 | ForEach-Object { $_ * 2 }) && NextStep
if ($log.Count -ne 1 -or $log[0] -ne "Next") {
    Write-Host "FAIL: && operator with pipeline on left failed"
    exit 1
}
Write-Host "PASS"
exit 0
