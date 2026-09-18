# vybe-test: powershell/engine_history_management/history_entry_commandline_property
# HistoryInfo exposes .CommandLine containing the exact command string
Clear-History
$cmdText = "Get-ChildItem -Path /tmp -Filter *.log"
$entry = [pscustomobject]@{
    CommandLine = $cmdText
    ExecutionStatus = "Completed"
    StartExecutionTime = [DateTime]::Now
    EndExecutionTime = [DateTime]::Now.AddMilliseconds(10)
}

$historyObj = Add-History -InputObject $entry -Passthru
Clear-History

if ($historyObj.CommandLine -ne $cmdText) {
    Write-Host "FAIL: CommandLine mismatch, expected '$cmdText', got: '$($historyObj.CommandLine)'"
    exit 1
}

Write-Host "PASS"
exit 0
