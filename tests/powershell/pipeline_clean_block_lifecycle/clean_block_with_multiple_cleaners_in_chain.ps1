# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_with_multiple_cleaners_in_chain
$global:C1 = $false
$global:C2 = $false
function Clean1 {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)]$In)
    process { $In }
    clean { $global:C1 = $true }
}
function Clean2 {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)]$In)
    process { $In }
    clean { $global:C2 = $true }
}
1, 2 | Clean1 | Clean2
if (-not $global:C1 -or -not $global:C2) {
    Write-Host "FAIL: Multiple cleaners in chain failed"
    exit 1
}
Write-Host "PASS"
exit 0
