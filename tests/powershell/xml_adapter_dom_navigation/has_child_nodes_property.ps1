# vybe-test: powershell/xml_adapter_dom_navigation/has_child_nodes_property
$xml = [xml]'<Hierarchy>
  <Branch><Leaf>Val</Leaf></Branch>
  <EmptyLeaf/>
</Hierarchy>'

# HasChildNodes returns true for nodes with children, false for empty nodes
$branchHasChildren = $xml.Hierarchy.Branch.HasChildNodes
$emptyHasChildren = $xml.Hierarchy.EmptyLeaf.HasChildNodes

if (-not $branchHasChildren) {
    Write-Host "FAIL: Branch.HasChildNodes expected `$true, got `$false"
    exit 1
}

if ($emptyHasChildren) {
    Write-Host "FAIL: EmptyLeaf.HasChildNodes expected `$false, got `$true"
    exit 1
}

Write-Host "PASS"
exit 0
