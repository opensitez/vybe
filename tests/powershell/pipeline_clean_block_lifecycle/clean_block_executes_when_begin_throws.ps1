# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_executes_when_begin_throws
$global:CleanBeginRan = $false
function Test-CleanBeginThrow {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    begin { throw "FailInBegin" }
    clean { $global:CleanBeginRan = $true }
}
$caught = $false
try {
    1, 2 | Test-CleanBeginThrow
} catch {
    $caught = $true
}
if (-not $caught -or -not $global:CleanBeginRan) {
    Write-Host "FAIL: Clean block should run when begin block throws"
    exit 1
}
Write-Host "PASS"
exit 0
