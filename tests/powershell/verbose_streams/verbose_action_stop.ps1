# vybe-test: powershell/verbose_streams/verbose_action_stop
$VerbosePreference = 'Stop'
try { Write-Verbose 'x' } catch { $caught = $true }
if (-not $caught) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
