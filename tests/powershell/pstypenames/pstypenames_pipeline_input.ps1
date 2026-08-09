# vybe-test: powershell/pstypenames/pstypenames_pipeline_input
$res = [pscustomobject]@{ Tag = "P1" } | ForEach-Object {
    $_.psobject.TypeNames.Insert(0, "PipelineType")
    $_.psobject.TypeNames[0]
}
if ($res -ne "PipelineType") {
    Write-Host "FAIL: TypeNames in pipeline expected PipelineType, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
