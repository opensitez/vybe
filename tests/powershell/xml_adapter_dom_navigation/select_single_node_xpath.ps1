# vybe-test: powershell/xml_adapter_dom_navigation/select_single_node_xpath
$xml = [xml]'<Employees>
  <Employee id="E100">Alice</Employee>
  <Employee id="E200">Bob</Employee>
</Employees>'

# SelectSingleNode returns exactly one XmlNode matching the XPath query
$node = $xml.SelectSingleNode('//Employee[@id="E200"]')

if ($null -eq $node) {
    Write-Host "FAIL: SelectSingleNode returned `$null"
    exit 1
}

if ($node.InnerText -ne "Bob") {
    Write-Host "FAIL: expected 'Bob', got '$($node.InnerText)'"
    exit 1
}

if ($node.Attributes["id"].Value -ne "E200") {
    Write-Host "FAIL: attribute id expected 'E200', got '$($node.Attributes['id'].Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
