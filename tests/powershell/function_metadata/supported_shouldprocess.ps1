# vybe-test: powershell/function_metadata/supported_shouldprocess
function Test-Func { [CmdletBinding(SupportsShouldProcess=$true)] param() if ($PSCmdlet.ShouldProcess('target')) { Write-Output 'PASS' } }
if ((Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
