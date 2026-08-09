# vybe-test: powershell/using_statements/using_assembly_system_xml
using namespace System.Xml

$doc = [XmlDocument]::new()
$doc.LoadXml("<x>y</x>")
if ($doc.DocumentElement.InnerText -ne "y") {
    Write-Host "FAIL: using namespace System.Xml XmlDocument expected 'y'"
    exit 1
}
Write-Host "PASS"
exit 0
