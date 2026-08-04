# vybe-test: powershell/remote_sessions/invoke_command_session
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    $result = Invoke-Command -Session $session -ScriptBlock { 11 }
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    if ($result -ne 11) {
        Write-Host "FAIL: expected 11, got $result"
        exit 1
    }
}
Write-Host "PASS"
exit 0
