# vybe-test: powershell/pipeline_begin_process_end_blocks/begin_block_throwing_prevents_process_execution
function Fail-InBegin {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    begin { throw "BeginFail" }
    process { "PROC" }
}
$caught = $false
try {
    1, 2, 3 | Fail-InBegin
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Exception in begin block should abort pipeline"
    exit 1
}
Write-Host "PASS"
exit 0
