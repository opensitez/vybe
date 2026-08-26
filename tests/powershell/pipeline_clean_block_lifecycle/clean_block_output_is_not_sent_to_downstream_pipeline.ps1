# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_output_is_not_sent_to_downstream_pipeline
function Test-CleanNoOutput {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process { $Val }
    clean { "CLEAN_SHOULD_NOT_POLLUTE" }
}
$res = @(1, 2 | Test-CleanNoOutput)
# In PowerShell 7.3+, clean block output does not emit to pipeline stream
if ($res.Length -ne 2 -or $res[0] -ne 1 -or $res[1] -ne 2) {
    Write-Host "FAIL: Clean block output emission behavior failed"
    exit 1
}
Write-Host "PASS"
exit 0
