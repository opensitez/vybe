# vybe-test: powershell/type_uri_parsing_and_components/combine_base_and_relative_uri
$base = [uri]"https://example.com/api/"
$rel = [uri]::new($base, "v2/users")
if ($rel.AbsoluteUri -ne "https://example.com/api/v2/users") {
    Write-Host "FAIL: Base and relative URI combine failed"
    exit 1
}
Write-Host "PASS"
exit 0
