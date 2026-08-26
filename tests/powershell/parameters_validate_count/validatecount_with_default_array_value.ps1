# vybe-test: powershell/parameters_validate_count/validatecount_with_default_array_value
function Get-ClusterNodes {
    param([ValidateCount(1, 3)][string[]]$Nodes = @("node1", "node2"))
    return $Nodes.Length
}
$res = Get-ClusterNodes
if ($res -ne 2) {
    Write-Host "FAIL: ValidateCount with default array failed"
    exit 1
}
Write-Host "PASS"
exit 0
