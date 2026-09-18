# vybe-test: powershell/xml_adapter_dom_navigation/outer_xml_includes_enclosing_tags
$xml = [xml]'<Container><Section><Item>DataPoint</Item></Section></Container>'

$outer = $xml.Container.Section.OuterXml
$inner = $xml.Container.Section.InnerXml

# OuterXml contains enclosing <Section>...</Section> tags
if ($outer -ne '<Section><Item>DataPoint</Item></Section>') {
    Write-Host "FAIL: OuterXml mismatch, got: '$outer'"
    exit 1
}

# InnerXml contains only the inner children
if ($inner -ne '<Item>DataPoint</Item>') {
    Write-Host "FAIL: InnerXml mismatch, got: '$inner'"
    exit 1
}

Write-Host "PASS"
exit 0
