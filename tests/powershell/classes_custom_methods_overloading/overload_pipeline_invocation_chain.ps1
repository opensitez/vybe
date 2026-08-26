# vybe-test: powershell/classes_custom_methods_overloading/overload_pipeline_invocation_chain
class PipeOverload {
    [int]Process([int]$x) { return $x * 2 }
}
$po = [PipeOverload]::new()
$res = @(1, 2, 3) | ForEach-Object { $po.Process($_) }
if ($res[0] -ne 2 -or $res[1] -ne 4 -or $res[2] -ne 6) {
    Write-Host "FAIL: Pipeline overload method invocation failed"
    exit 1
}
Write-Host "PASS"
exit 0
