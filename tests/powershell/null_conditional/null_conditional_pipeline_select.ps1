# vybe-test: powershell/null_conditional/null_conditional_pipeline_select
$list = @([pscustomobject]@{ Code = "C1" }, $null)
$res = $list | ForEach-Object { ${_}?.Code }
if ($res[0] -ne "C1" -or $res[1] -ne $null) {
    Write-Host "FAIL: pipeline null-conditional expected C1, null"
    exit 1
}
Write-Host "PASS"
exit 0
