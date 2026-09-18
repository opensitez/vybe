# vybe-test: powershell/xml_adapter_dom_navigation/cdata_section_inner_text_preservation
$xml = [xml]'<Template><Script><![CDATA[if (a < 10 && b > 20) { return "ok"; }]]></Script></Template>'

# CDATA section preserves raw unescaped characters without XML entity parsing
$scriptText = $xml.Template.Script.InnerText

$expected = 'if (a < 10 && b > 20) { return "ok"; }'

if ($scriptText -ne $expected) {
    Write-Host "FAIL: CDATA text mismatch, expected '$expected', got '$scriptText'"
    exit 1
}

Write-Host "PASS"
exit 0
