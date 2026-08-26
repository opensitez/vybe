# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_with_select_object_cmdlet
$p = @{ First = 2 }
$res = @(1..10 | Select-Object @p)
if ($res.Length -ne 2 -or $res[1] -ne 2) {
    Write-Host "FAIL: Splatting with standard cmdlet Select-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
