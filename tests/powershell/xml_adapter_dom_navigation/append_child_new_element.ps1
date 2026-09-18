# vybe-test: powershell/xml_adapter_dom_navigation/append_child_new_element
$xml = [xml]'<Folder id="root"/>'

# Dynamically create and append an element to DocumentElement
$newElem = $xml.CreateElement("File")
$newElem.InnerText = "readme.txt"
$newElem.SetAttribute("size", "1024")

$xml.DocumentElement.AppendChild($newElem) | Out-Null

# Verify the appended child is accessible via dot notation
$fileName = $xml.DocumentElement.File.InnerText
$fileSize = $xml.DocumentElement.File.size

if ($fileName -ne "readme.txt" -or $fileSize -ne "1024") {
    Write-Host "FAIL: appended element not correctly accessible via dot syntax, got name='$fileName', size='$fileSize'"
    exit 1
}

Write-Host "PASS"
exit 0
