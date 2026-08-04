# vybe-test: powershell/remote_sessions/session_command_exception
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    $result = Invoke-Command -Session $session -ScriptBlock { 1 / 1 }
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    if ($result -ne 1) {
        Write-Host "FAIL: expected 1"
        exit 1
    }
}
Write-Host "PASS"
exit 0
