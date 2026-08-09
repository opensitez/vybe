# vybe-test: powershell/using_statements/using_namespace_net
using namespace System.Net

$ep = [IPEndPoint]::new([IPAddress]::Loopback, 8080)
if ($ep.Port -ne 8080) {
    Write-Host "FAIL: IPEndPoint Port expected 8080, got $($ep.Port)"
    exit 1
}
Write-Host "PASS"
exit 0
