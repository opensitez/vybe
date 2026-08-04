# vybe-test: powershell/implicit_remoting/imported_command_available
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    Import-PSSession -Session $session -Module Microsoft.PowerShell.Core -ErrorAction SilentlyContinue | Out-Null
    if (-not (Get-Command Get-Process -ErrorAction SilentlyContinue)) {
        Write-Host "FAIL: expected imported command available"
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
        exit 1
    }
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
}
Write-Host "PASS"
exit 0
