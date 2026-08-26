# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_inside_scriptblock
$log = [System.Collections.Generic.List[string]]::new()
$sb = {
    param([string]$Tag)
    $log.Add($Tag)
    return $true
}
(& $sb "First") && (& $sb "Second")
if ($log.Count -ne 2 -or $log[0] -ne "First" -or $log[1] -ne "Second") {
    Write-Host "FAIL: && inside scriptblock invocation failed"
    exit 1
}
Write-Host "PASS"
exit 0
