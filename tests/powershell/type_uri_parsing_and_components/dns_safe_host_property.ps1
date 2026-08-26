# vybe-test: powershell/type_uri_parsing_and_components/dns_safe_host_property
$u = [uri]"https://my-host.sub.domain.com:9000/"
if ($u.DnsSafeHost -ne "my-host.sub.domain.com") {
    Write-Host "FAIL: DnsSafeHost extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
