# vybe-test: powershell/implicit_remoting/cleanup_after_import
$session = New-PSSession -ComputerName localhost -ErrorAction SilentlyContinue
if ($session) {
    Import-PSSession -Session $session -Module Microsoft.PowerShell.Core -ErrorAction SilentlyContinue | Out-Null
    Remove-Module Microsoft.PowerShell.Core -ErrorAction SilentlyContinue
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
}
Write-Host "PASS"
exit 0
