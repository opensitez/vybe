# vybe-test: powershell/xml_adapter_dom_navigation/select_xml_xpath_position_filter
$xml = [xml]'<Queue>
  <Message seq="1">FirstMsg</Message>
  <Message seq="2">MiddleMsg</Message>
  <Message seq="3">LastMsg</Message>
</Queue>'

# XPath positional functions: last() matches the final node in sequence
$lastMatch = Select-Xml -Xml $xml -XPath '//Message[last()]'

if ($null -eq $lastMatch -or $lastMatch.Node.InnerText -ne "LastMsg") {
    Write-Host "FAIL: XPath [last()] failed, expected 'LastMsg', got '$($lastMatch.Node.InnerText)'"
    exit 1
}

# XPath index 1 matches the first node in 1-based XPath indexing
$firstMatch = Select-Xml -Xml $xml -XPath '//Message[1]'
if ($firstMatch.Node.InnerText -ne "FirstMsg") {
    Write-Host "FAIL: XPath [1] failed, expected 'FirstMsg', got '$($firstMatch.Node.InnerText)'"
    exit 1
}

Write-Host "PASS"
exit 0
