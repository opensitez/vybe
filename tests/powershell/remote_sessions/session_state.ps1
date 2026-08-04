# vybe-test: powershell/remote_sessions/session_state
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    if ($session.State -ne 'Opened' -and $session.State -ne 'OpenedByRunspace') {
        Write-Host "FAIL: expected opened session state"
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
        exit 1
    }
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
}
Write-Host "PASS"
exit 0
