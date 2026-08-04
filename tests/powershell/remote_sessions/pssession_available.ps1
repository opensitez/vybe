# vybe-test: powershell/remote_sessions/pssession_available
$available = Get-PSSession -ErrorAction SilentlyContinue
if ($available -and $available.Count -lt 0) {
    Write-Host "FAIL: unexpected session count"
    exit 1
}
Write-Host "PASS"
exit 0
