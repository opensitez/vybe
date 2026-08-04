# vybe-test: powershell/variable_modifiers/read_only_error
Set-Variable -Name x -Value 1 -Option ReadOnly
$failed = $false
try { $script:x = 2 } catch { $failed = $true }
if ($failed -and $x -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
