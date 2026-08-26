# vybe-test: powershell/pipeline_begin_process_end_blocks/end_block_returns_array_of_collected_items
function Reverse-Pipeline {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    begin { $list = [System.Collections.Generic.List[int]]::new() }
    process { $list.Add($Val) }
    end {
        $list.Reverse()
        return $list.ToArray()
    }
}
$res = 1, 2, 3, 4 | Reverse-Pipeline
if ($res.Length -ne 4 -or $res[0] -ne 4 -or $res[3] -ne 1) {
    Write-Host "FAIL: Reverse pipeline in end block failed"
    exit 1
}
Write-Host "PASS"
exit 0
