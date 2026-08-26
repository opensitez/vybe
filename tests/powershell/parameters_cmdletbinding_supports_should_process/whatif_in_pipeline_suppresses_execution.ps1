# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/whatif_in_pipeline_suppresses_execution
function Remove-PipeItem2 {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true)]
        [string]$Name
    )
    process {
        if ($PSCmdlet.ShouldProcess($Name, "Remove")) {
            "Removed:$Name"
        }
    }
}
$res = @("Item1", "Item2" | Remove-PipeItem2 -WhatIf)
if ($res.Count -ne 0) {
    Write-Host "FAIL: -WhatIf in pipeline should suppress all process block body executions"
    exit 1
}
Write-Host "PASS"
exit 0
