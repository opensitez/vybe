# vybe-test: powershell/dsc_lcm/lcm_describe
$cmd = Get-Command Get-DscLocalConfigurationManager -ErrorAction SilentlyContinue
if (-not $cmd) { Write-Host 'FAIL' ; exit 1 }
Write-Host 'PASS'
exit 0
