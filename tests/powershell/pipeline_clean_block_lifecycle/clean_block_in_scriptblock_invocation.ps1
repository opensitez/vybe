# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_in_scriptblock_invocation
$global:SbClean = $false
$sb = {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process { $Val * 2 }
    clean { $global:SbClean = $true }
}
$res = 5 | & $sb
if ($res -ne 10 -or -not $global:SbClean) {
    Write-Host "FAIL: Clean block in scriptblock execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
