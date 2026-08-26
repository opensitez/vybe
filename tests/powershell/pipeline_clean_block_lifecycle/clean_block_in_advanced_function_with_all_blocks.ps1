# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_in_advanced_function_with_all_blocks
function Full-LifecycleFunction {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    begin { $order = "B" }
    process { $order += "P" }
    end { $order += "E" }
    clean { $order += "C"; return $order }
}
$res = 1 | Full-LifecycleFunction
if (-not $res.Contains("BPE")) {
    Write-Host "FAIL: Full lifecycle function check failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
