# vybe-test: powershell/engine_history_management/history_multiline_commandline_preservation
# Multiline commands preserve embedded newline characters within HistoryInfo.CommandLine
Clear-History
$multiline = "Get-Service |`n  Where-Object Status -eq 'Running' |`n  Select-Object -First 5"
$entry = Add-History -InputObject ([pscustomobject]@{
    CommandLine = $multiline
    ExecutionStatus = "Completed"
    StartExecutionTime = [DateTime]::Now
    EndExecutionTime = [DateTime]::Now
}) -Passthru
Clear-History

if ($entry.CommandLine -ne $multiline) {
    Write-Host "FAIL: multiline command formatting was altered: '$($entry.CommandLine)'"
    exit 1
}

Write-Host "PASS"
exit 0
