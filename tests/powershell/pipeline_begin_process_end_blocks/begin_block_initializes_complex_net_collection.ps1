# vybe-test: powershell/pipeline_begin_process_end_blocks/begin_block_initializes_complex_net_collection
function Collect-HashList {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][string]$Key)
    begin { $set = [System.Collections.Generic.HashSet[string]]::new() }
    process { $null = $set.Add($Key) }
    end { return $set.Count }
}
$res = "a", "b", "a", "c", "b" | Collect-HashList
if ($res -ne 3) {
    Write-Host "FAIL: Complex collection in begin block failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
