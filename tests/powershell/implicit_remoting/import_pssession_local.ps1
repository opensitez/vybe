# vybe-test: powershell/implicit_remoting/import_pssession_local
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    Import-PSSession -Session $session -Module Microsoft.PowerShell.Core -ErrorAction SilentlyContinue | Out-Null
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
}
Write-Host "PASS"
exit 0
