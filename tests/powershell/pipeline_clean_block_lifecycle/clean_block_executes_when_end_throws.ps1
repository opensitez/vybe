# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_executes_when_end_throws
$global:CleanEndRan = $false
function Test-CleanEndThrow {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    end { throw "FailInEnd" }
    clean { $global:CleanEndRan = $true }
}
$caught = $false
try {
    1, 2 | Test-CleanEndThrow
} catch {
    $caught = $true
}
if (-not $caught -or -not $global:CleanEndRan) {
    Write-Host "FAIL: Clean block should run when end block throws"
    exit 1
}
Write-Host "PASS"
exit 0
