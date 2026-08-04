# vybe-test: powershell/remote_sessions/session_location
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    if ($session.ComputerName -ne 'localhost') {
        Write-Host "FAIL: expected localhost session"
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
        exit 1
    }
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
}
Write-Host "PASS"
exit 0
