# vybe-test: powershell/type_uri_parsing_and_components/query_parameter_split
$u = [uri]"https://example.com/test?a=1&b=hello"
$q = $u.Query.TrimStart('?')
$pairs = $q -split '&'
if ($pairs.Length -ne 2 -or $pairs[0] -ne "a=1" -or $pairs[1] -ne "b=hello") {
    Write-Host "FAIL: Query string manual split failed"
    exit 1
}
Write-Host "PASS"
exit 0
