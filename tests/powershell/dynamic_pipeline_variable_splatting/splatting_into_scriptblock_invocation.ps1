# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_into_scriptblock_invocation
function Invoke-SplatTarget {
    [CmdletBinding()]
    param([string]$Prefix, [string]$Suffix)
    return "${Prefix}:Data:${Suffix}"
}
$p = @{ Prefix = "START"; Suffix = "END" }
$res = Invoke-SplatTarget @p
if ($res -ne "START:Data:END") {
    Write-Host "FAIL: Function invocation failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
