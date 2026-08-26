# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_with_pipeline_on_right
$res = [System.Collections.Generic.List[int]]::new()
function SuccessCmd { return $true }
SuccessCmd && (1..3 | ForEach-Object { $res.Add($_) })
if ($res.Count -ne 3 -or $res[0] -ne 1 -or $res[2] -ne 3) {
    Write-Host "FAIL: && operator with pipeline on right failed"
    exit 1
}
Write-Host "PASS"
exit 0
