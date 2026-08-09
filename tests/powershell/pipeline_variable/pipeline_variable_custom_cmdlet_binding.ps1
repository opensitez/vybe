# vybe-test: powershell/pipeline_variable/pipeline_variable_custom_cmdlet_binding
function Test-PVBinding {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [int]$InputObject
    )
    process {
        return $InputObject * 3
    }
}
$res = 1..2 | Test-PVBinding -PipelineVariable pv | ForEach-Object { "$pv:$_" }
if ($res[0] -ne "1:3" -or $res[1] -ne "2:6") {
    Write-Host "FAIL: custom CmdletBinding function -PipelineVariable expected 1:3, 2:6, got $($res -join ', ')"
    exit 1
}
Write-Host "PASS"
exit 0
