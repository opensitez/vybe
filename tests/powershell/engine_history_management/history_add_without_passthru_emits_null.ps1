# vybe-test: powershell/engine_history_management/history_add_without_passthru_emits_null
# Calling Add-History without -Passthru outputs nothing ($null)
Clear-History
$out = Add-History -InputObject ([pscustomobject]@{
    CommandLine = "SilentAdd"
    ExecutionStatus = "Completed"
    StartExecutionTime = [DateTime]::Now
    EndExecutionTime = [DateTime]::Now
})
Clear-History

if ($null -ne $out) {
    Write-Host "FAIL: Add-History without -Passthru unexpectedly emitted: $out"
    exit 1
}

Write-Host "PASS"
exit 0
