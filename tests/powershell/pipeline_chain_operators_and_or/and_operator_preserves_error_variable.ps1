# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_preserves_error_variable
$errs = $null
function MakeError { Write-Error "SpecificError" }
function AfterError { return $true }
MakeError && AfterError
if ($Error.Count -eq 0 -or -not $Error[0].ToString().Contains("SpecificError")) {
    Write-Host "FAIL: Error record preservation failed"
    exit 1
}
Write-Host "PASS"
exit 0
