# vybe-test: powershell/aliasing/alias_with_command_info
Set-Alias hi Write-Output
$cmd = Get-Command hi
if ($cmd.CommandType -ne 'Alias') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
