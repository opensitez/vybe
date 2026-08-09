# vybe-test: powershell/pipeline_chaining/chain_in_function
function Test-Chain([bool]$a, [bool]$b) {
    return ($a) && ($b)
}
$res1 = Test-Chain $true $true
$res2 = Test-Chain $true $false
if (-not $res1 -or $res2) {
    Write-Host "FAIL: function pipeline chaining expected res1=true, res2=false"
    exit 1
}
Write-Host "PASS"
exit 0
