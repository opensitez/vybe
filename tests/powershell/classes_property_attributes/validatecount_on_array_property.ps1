# vybe-test: powershell/classes_property_attributes/validatecount_on_array_property
class Cluster {
    [ValidateCount(1, 3)][string[]]$Nodes
}
$c = [Cluster]::new()
$c.Nodes = @("Node1", "Node2")
$caught = $false
try {
    $c.Nodes = @("N1", "N2", "N3", "N4")
} catch {
    $caught = $true
}
if ($c.Nodes.Length -ne 2 -or -not $caught) {
    Write-Host "FAIL: ValidateCount on array property failed"
    exit 1
}
Write-Host "PASS"
exit 0
