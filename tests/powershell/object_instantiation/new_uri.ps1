# vybe-test: powershell/object_instantiation/new_uri
$uri = [uri]'https://example.com'
if ($uri.Host -ne 'example.com') { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
