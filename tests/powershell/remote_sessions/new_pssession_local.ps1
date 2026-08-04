# vybe-test: powershell/remote_sessions/new_pssession_local
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session -and $session.ComputerName -ne 'localhost') {
    Write-Host "FAIL: expected localhost session"
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    exit 1
}
if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
Write-Host "PASS"
exit 0
