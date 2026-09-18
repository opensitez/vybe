# vybe-test: powershell/engine_history_management/history_clear_by_commandline_wildcard
# Clear-History -CommandLine supports wildcard patterns to batch remove matching commands
Clear-History
Add-History -InputObject ([pscustomobject]@{ CommandLine = "Get-Process -Name pwsh"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })
Add-History -InputObject ([pscustomobject]@{ CommandLine = "Get-Service sshd"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })
Add-History -InputObject ([pscustomobject]@{ CommandLine = "Set-Location /"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })

Clear-History -CommandLine "Get-*"

$remaining = @(Get-History | ForEach-Object { $_.CommandLine })
Clear-History

if ($remaining -contains "Get-Process -Name pwsh" -or $remaining -contains "Get-Service sshd") {
    Write-Host "FAIL: commands matching 'Get-*' were not removed: @($($remaining -join ', '))"
    exit 1
}

if ($remaining -notcontains "Set-Location /") {
    Write-Host "FAIL: 'Set-Location /' was unexpectedly removed"
    exit 1
}

Write-Host "PASS"
exit 0
