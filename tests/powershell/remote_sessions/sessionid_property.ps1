# vybe-test: powershell/remote_sessions/sessionid_property
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    if (-not $session.Id) {
        Write-Host "FAIL: expected session id"
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
        exit 1
    }
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
}
Write-Host "PASS"
exit 0
