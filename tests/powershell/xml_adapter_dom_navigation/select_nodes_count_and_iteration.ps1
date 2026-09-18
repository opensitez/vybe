# vybe-test: powershell/xml_adapter_dom_navigation/select_nodes_count_and_iteration
$xml = [xml]'<Project>
  <Task priority="high">Design</Task>
  <Task priority="medium">Build</Task>
  <Task priority="high">Test</Task>
</Project>'

# SelectNodes returns an XmlNodeList collection
$nodes = $xml.SelectNodes('//Task[@priority="high"]')

if ($nodes.Count -ne 2) {
    Write-Host "FAIL: expected 2 high priority tasks, got $($nodes.Count)"
    exit 1
}

$taskNames = @()
foreach ($n in $nodes) {
    $taskNames += $n.InnerText
}

if ($taskNames[0] -ne "Design" -or $taskNames[1] -ne "Test") {
    Write-Host "FAIL: task iteration expected @('Design', 'Test'), got @($($taskNames -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
