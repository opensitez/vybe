# vybe-test: powershell/function_scope/loop_scope
# A loop body is NOT a scope — `$x` assigned inside the `for` is visible after
# it. The enclosing FUNCTION is the scope, so that same `$x` does not escape to
# the caller.
function Test-Func {
    for ($i = 0; $i -lt 1; $i++) { $x = 2 }
    return $x
}
$inner = Test-Func
if ($inner -ne 2) { Write-Host "FAIL: loop body did not promote x, got $inner"; exit 1 }
if ($null -ne $x) { Write-Host "FAIL: x escaped the function as $x"; exit 1 }
Write-Host 'PASS'
exit 0
