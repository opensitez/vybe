# vybe-test: powershell/xml_adapter_dom_navigation/select_xml_xpath_attribute_filter
$xml = [xml]'<Cluster>
  <Node role="primary" ip="10.0.0.1">MasterNode</Node>
  <Node role="worker" ip="10.0.0.2">WorkerOne</Node>
  <Node role="worker" ip="10.0.0.3">WorkerTwo</Node>
</Cluster>'

# Select-Xml queries nodes using XPath expressions filtering by attribute
$matched = Select-Xml -Xml $xml -XPath '//Node[@role="worker"]'

if ($null -eq $matched) {
    Write-Host "FAIL: Select-Xml returned `$null"
    exit 1
}

$nodes = @($matched | ForEach-Object { $_.Node.InnerText })

if ($nodes.Count -ne 2 -or $nodes[0] -ne "WorkerOne" -or $nodes[1] -ne "WorkerTwo") {
    Write-Host "FAIL: expected worker nodes @('WorkerOne', 'WorkerTwo'), got @($($nodes -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
