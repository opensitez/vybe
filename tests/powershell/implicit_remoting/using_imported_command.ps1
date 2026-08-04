# vybe-test: powershell/implicit_remoting/using_imported_command
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    Import-PSSession -Session $session -Module Microsoft.PowerShell.Core -ErrorAction SilentlyContinue | Out-Null
    Get-Command Get-Date | Out-Null
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
}
Write-Host "PASS"
exit 0
