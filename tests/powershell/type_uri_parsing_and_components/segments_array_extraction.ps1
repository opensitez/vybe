# vybe-test: powershell/type_uri_parsing_and_components/segments_array_extraction
$u = [uri]"https://example.com/a/b/c/"
$seg = $u.Segments
if ($seg.Length -ne 4 -or $seg[1] -ne "a/" -or $seg[2] -ne "b/" -or $seg[3] -ne "c/") {
    Write-Host "FAIL: Segments array extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
