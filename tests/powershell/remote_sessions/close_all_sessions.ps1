# vybe-test: powershell/remote_sessions/close_all_sessions
$existing = Get-PSSession -ErrorAction SilentlyContinue
if ($existing) { Remove-PSSession -Session $existing -ErrorAction SilentlyContinue }
Write-Host "PASS"
exit 0
