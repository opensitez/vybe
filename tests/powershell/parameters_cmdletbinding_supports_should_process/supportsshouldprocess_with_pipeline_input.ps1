# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/supportsshouldprocess_with_pipeline_input
function Remove-PipeItem {
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
$res = "Item1", "Item2" | Remove-PipeItem
if ($res.Length -ne 2 -or $res[0] -ne "Removed:Item1" -or $res[1] -ne "Removed:Item2") {
    Write-Host "FAIL: SupportsShouldProcess with pipeline input failed"
    exit 1
}
Write-Host "PASS"
exit 0
