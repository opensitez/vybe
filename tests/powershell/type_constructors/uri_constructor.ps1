# vybe-test: powershell/type_constructors/uri_constructor
$value = [uri]'https://example.com'
if ($value.Host -ne 'example.com') { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
