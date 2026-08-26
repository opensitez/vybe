# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_accesses_begin_block_variables
$global:RecordedBeginVal = ""
function Test-VarAccess {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][string]$Val)
    begin { $msg = "Initialized" }
    clean { $global:RecordedBeginVal = $msg }
}
"test" | Test-VarAccess
if ($global:RecordedBeginVal -ne "Initialized") {
    Write-Host "FAIL: Clean block accessing begin block variable failed"
    exit 1
}
Write-Host "PASS"
exit 0
