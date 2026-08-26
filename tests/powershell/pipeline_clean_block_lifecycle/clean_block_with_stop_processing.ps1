# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_with_stop_processing
$global:CleanStopRan = $false
function Test-SelectFirstClean {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process { $Val }
    clean { $global:CleanStopRan = $true }
}
$res = 1..100 | Test-SelectFirstClean | Select-Object -First 1
if ($res -ne 1 -or -not $global:CleanStopRan) {
    Write-Host "FAIL: Clean block with Select-Object -First short-circuit failed"
    exit 1
}
Write-Host "PASS"
exit 0
