# vybe-test: powershell/pipeline_chaining/chain_pipeline_output
$res = 1..3 | ForEach-Object { ($_ -gt 1) && "Gt1" }
if ($res[0] -ne $null -or $res[1] -ne "Gt1" -or $res[2] -ne "Gt1") {
    Write-Host "FAIL: pipeline element chaining expected null, Gt1, Gt1"
    exit 1
}
Write-Host "PASS"
exit 0
