# vybe-test: powershell/script_parameters/pipeline_parameter
param([Parameter(ValueFromPipeline=$true)]$x)
process { if ($x -eq 1) { Write-Host 'PASS'; exit 0 } }
Write-Host 'FAIL'
exit 1
