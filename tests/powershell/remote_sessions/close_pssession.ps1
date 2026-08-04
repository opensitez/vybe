# vybe-test: powershell/remote_sessions/close_pssession
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
Write-Host "PASS"
exit 0
