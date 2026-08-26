# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_executes_on_normal_completion
function Test-CleanSuccess {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    begin { $events = [System.Collections.Generic.List[string]]::new() }
    process { $events.Add("P:$Val") }
    end { $events.Add("END") }
    clean { $events.Add("CLEAN"); $events -join ";" }
}
$res = 1, 2 | Test-CleanSuccess
if (-not $res.Contains("CLEAN") -or -not $res.Contains("END")) {
    Write-Host "FAIL: Clean block normal completion check failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
