# vybe-test: powershell/variable_modifiers/constant
Set-Variable -Name x -Value 1 -Option Constant
$failed = $false
try { Set-Variable -Name x -Value 2 } catch { $failed = $true }
if ($failed -and $x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
