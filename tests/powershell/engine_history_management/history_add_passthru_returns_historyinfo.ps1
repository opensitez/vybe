# vybe-test: powershell/engine_history_management/history_add_passthru_returns_historyinfo
# Add-History -Passthru emits instances of Microsoft.PowerShell.Commands.HistoryInfo
Clear-History
$entry = [pscustomobject]@{
    CommandLine = "Get-Process -Name pwsh"
    ExecutionStatus = "Completed"
    StartExecutionTime = [DateTime]::Now
    EndExecutionTime = [DateTime]::Now.AddMilliseconds(25)
}

$historyObj = Add-History -InputObject $entry -Passthru
Clear-History

if ($null -eq $historyObj) {
    Write-Host "FAIL: Add-History -Passthru returned `$null"
    exit 1
}

if (-not ($historyObj -is [Microsoft.PowerShell.Commands.HistoryInfo])) {
    Write-Host "FAIL: expected HistoryInfo type, got: $($historyObj.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
