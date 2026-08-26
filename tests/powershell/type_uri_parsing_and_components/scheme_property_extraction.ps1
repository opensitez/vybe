# vybe-test: powershell/type_uri_parsing_and_components/scheme_property_extraction
$ftp = [uri]"ftp://files.example.com/pub/doc.pdf"
if ($ftp.Scheme -ne "ftp") {
    Write-Host "FAIL: Scheme extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
