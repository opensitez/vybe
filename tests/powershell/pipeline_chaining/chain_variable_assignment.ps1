# vybe-test: powershell/pipeline_chaining/chain_variable_assignment
$outVal = ($true) && "AssignedValue"
if ($outVal -ne "AssignedValue") {
    Write-Host "FAIL: chain variable assignment expected AssignedValue, got $outVal"
    exit 1
}
Write-Host "PASS"
exit 0
