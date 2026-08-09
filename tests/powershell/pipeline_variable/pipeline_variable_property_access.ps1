# vybe-test: powershell/pipeline_variable/pipeline_variable_property_access
$objs = @([pscustomobject]@{ Code = "A" }, [pscustomobject]@{ Code = "B" })
$res = $objs | ForEach-Object -PipelineVariable o { $_.Code } | ForEach-Object { "$($o.Code):$_" }
if ($res[0] -ne "A:A" -or $res[1] -ne "B:B") {
    Write-Host "FAIL: PipelineVariable property access expected A:A, B:B"
    exit 1
}
Write-Host "PASS"
exit 0
