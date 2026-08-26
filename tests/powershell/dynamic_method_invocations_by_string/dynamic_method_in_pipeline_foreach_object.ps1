# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_in_pipeline_foreach_object
$methods = @("ToLower", "ToUpper")
$str = "MixedCase"
$res = @($methods | ForEach-Object { $str.$_() })
if ($res[0] -ne "mixedcase" -or $res[1] -ne "MIXEDCASE") {
    Write-Host "FAIL: Dynamic method in pipeline ForEach-Object failed"
    exit 1
}
Write-Host "PASS"
exit 0
