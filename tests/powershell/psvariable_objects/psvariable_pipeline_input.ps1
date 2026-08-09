# vybe-test: powershell/psvariable_objects/psvariable_pipeline_input
$res = Get-Variable -Name "PID", "PWD"
if ($res.Count -ne 2) {
    Write-Host "FAIL: Get-Variable multiple names expected 2 objects"
    exit 1
}
Write-Host "PASS"
exit 0
