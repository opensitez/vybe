# vybe-test: powershell/pipeline_begin_process_end_blocks/process_block_throwing_terminates_pipeline
function Fail-OnTwo {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process {
        if ($Val -eq 2) { throw "Encountered 2" }
        $Val
    }
}
$caught = $false
$collected = [System.Collections.Generic.List[int]]::new()
try {
    1, 2, 3 | Fail-OnTwo | ForEach-Object { $collected.Add($_) }
} catch {
    $caught = $true
}
if (-not $caught -or $collected.Count -ne 1 -or $collected[0] -ne 1) {
    Write-Host "FAIL: Exception in process block should terminate pipeline"
    exit 1
}
Write-Host "PASS"
exit 0
