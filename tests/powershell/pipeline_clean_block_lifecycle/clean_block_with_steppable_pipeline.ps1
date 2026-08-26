# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_with_steppable_pipeline
function Test-StepClean {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process { $Val }
    clean { $global:StepCleanRan = $true }
}
$global:StepCleanRan = $false
$sp = { Test-StepClean }.GetSteppablePipeline()
$sp.Begin($true)
$sp.Process(10)
$sp.End()
$sp.Clean()
if (-not $global:StepCleanRan) {
    Write-Host "FAIL: SteppablePipeline.Clean() invocation failed"
    exit 1
}
Write-Host "PASS"
exit 0
