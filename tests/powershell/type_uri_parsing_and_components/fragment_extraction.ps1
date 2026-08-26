# vybe-test: powershell/type_uri_parsing_and_components/fragment_extraction
$u = [uri]"https://example.com/doc.html#section2"
if ($u.Fragment -ne "#section2") {
    Write-Host "FAIL: Fragment extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
