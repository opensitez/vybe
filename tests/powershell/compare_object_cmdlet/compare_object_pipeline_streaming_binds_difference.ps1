# vybe-test: powershell/compare_object_cmdlet/compare_object_pipeline_streaming_binds_difference
# In PowerShell, pipeline input to Compare-Object binds to -DifferenceObject (not ReferenceObject)
$diff = @(2, 3, 4) | Compare-Object -ReferenceObject @(1, 2, 3)

if ($diff.Count -ne 2) {
    Write-Host "FAIL: expected 2 diff records from pipeline, got $($diff.Count)"
    exit 1
}

$refOnly  = $diff | Where-Object { $_.SideIndicator -eq '<=' }
$diffOnly = $diff | Where-Object { $_.SideIndicator -eq '=>' }

if ($refOnly.InputObject -ne 1 -or $diffOnly.InputObject -ne 4) {
    Write-Host "FAIL: pipeline binding mismatch, refOnly=$($refOnly.InputObject), diffOnly=$($diffOnly.InputObject)"
    exit 1
}

Write-Host "PASS"
exit 0
