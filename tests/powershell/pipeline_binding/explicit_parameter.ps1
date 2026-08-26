# vybe-test: powershell/pipeline_binding/explicit_parameter
function Test-ExplicitBinding {
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process { return $Val * 2 }
}
$res = 21 | Test-ExplicitBinding
if ($res -ne 42) {
    Write-Host "FAIL: Explicit parameter binding failed"
    exit 1
}
Write-Host "PASS"
exit 0
