# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_disposes_resources
function Test-CleanFeature {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$InputObject)
    begin { $cleaned = $false }
    process { $InputObject * 2 }
    end { $cleaned = $true }
}
$res = @(1, 2, 3 | Test-CleanFeature)
if ($res.Length -ne 3 -or $res[0] -ne 2 -or $res[2] -ne 6) {
    Write-Host "FAIL: Pipeline clean feature execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
