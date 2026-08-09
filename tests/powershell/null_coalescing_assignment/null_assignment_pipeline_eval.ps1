# vybe-test: powershell/null_coalescing_assignment/null_assignment_pipeline_eval
$res = @($null, "Existing") | ForEach-Object {
    $x = $_
    $x ??= "Fallback"
    $x
}
if ($res[0] -ne "Fallback" -or $res[1] -ne "Existing") {
    Write-Host "FAIL: pipeline ??= expected Fallback, Existing"
    exit 1
}
Write-Host "PASS"
exit 0
