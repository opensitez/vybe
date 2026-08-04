# vybe-test: powershell/remote_sessions/session_command_output
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    $value = Invoke-Command -Session $session -ScriptBlock { 'session' }
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    if ($value -ne 'session') {
        Write-Host "FAIL: expected session output"
        exit 1
    }
}
Write-Host "PASS"
exit 0
