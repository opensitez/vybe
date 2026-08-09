# vybe-test: powershell/type_accelerators/type_accelerator_uri
$u = [uri]"https://api.github.com:443/repos/owner/repo?query=1#section"
if ($u.Scheme -ne "https") {
    Write-Host "FAIL: Scheme expected https, got $($u.Scheme)"
    exit 1
}
if ($u.Host -ne "api.github.com") {
    Write-Host "FAIL: Host expected api.github.com, got $($u.Host)"
    exit 1
}
if ($u.Port -ne 443) {
    Write-Host "FAIL: Port expected 443, got $($u.Port)"
    exit 1
}
if ($u.Query -ne "?query=1") {
    Write-Host "FAIL: Query expected ?query=1, got $($u.Query)"
    exit 1
}
Write-Host "PASS"
exit 0
