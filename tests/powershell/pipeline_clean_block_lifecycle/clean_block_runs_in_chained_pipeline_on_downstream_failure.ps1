# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_runs_in_chained_pipeline_on_downstream_failure
$global:UpstreamClean = $false
function Upstream-Producer {
    [CmdletBinding()]
    param()
    process { 1; 2; 3 }
    clean { $global:UpstreamClean = $true }
}
function Downstream-Failer {
    param([Parameter(ValueFromPipeline=$true)]$In)
    process { throw "DownstreamFail" }
}
try {
    Upstream-Producer | Downstream-Failer
} catch {}
if (-not $global:UpstreamClean) {
    Write-Host "FAIL: Upstream clean block should execute on downstream pipeline error"
    exit 1
}
Write-Host "PASS"
exit 0
