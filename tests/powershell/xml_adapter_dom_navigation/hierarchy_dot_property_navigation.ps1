# vybe-test: powershell/xml_adapter_dom_navigation/hierarchy_dot_property_navigation
$xml = [xml]'<Company><Department><Team><Lead>Sarah</Lead></Team></Department></Company>'

# PowerShell XML adapter enables multi-level hierarchical dot navigation
$lead = $xml.Company.Department.Team.Lead

if ($lead -ne "Sarah") {
    Write-Host "FAIL: expected 'Sarah', got: '$lead'"
    exit 1
}

Write-Host "PASS"
exit 0
