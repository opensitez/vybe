# vybe-test: powershell/pipeline_variable/pipeline_variable_group_object
$data = @([pscustomobject]@{ Type = "A"; Val = 1 }, [pscustomobject]@{ Type = "A"; Val = 2 })
$res = $data | Group-Object Type -PipelineVariable grp | ForEach-Object { "$($grp.Name):$($grp.Count)" }
if ($res -ne "A:2") {
    Write-Host "FAIL: Group-Object -PipelineVariable expected A:2, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
