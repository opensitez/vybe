# vybe-test: powershell/pipeline_begin_process_end_blocks/pipeline_with_multiple_named_parameters
function Filter-Threshold {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)][int]$Num,
        [Parameter()][int]$Min = 0
    )
    process {
        if ($Num -ge $Min) { $Num }
    }
}
$res = @(1, 5, 10, 2, 8 | Filter-Threshold -Min 5)
if ($res.Length -ne 3 -or $res[0] -ne 5 -or $res[1] -ne 10 -or $res[2] -ne 8) {
    Write-Host "FAIL: Named parameter along with pipeline input failed"
    exit 1
}
Write-Host "PASS"
exit 0
