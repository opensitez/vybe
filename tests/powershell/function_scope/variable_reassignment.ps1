# vybe-test: powershell/function_scope/variable_reassignment
# Assigning a name inside a function REASSIGNS it only within that function.
# PowerShell creates the binding in the local scope, so the caller's `$x` is
# untouched; writing the caller's storage needs `$script:x` / `$global:x`.
$x = 1
function Test-Func { $x = 2; return $x }
$inner = Test-Func
if ($inner -ne 2) { Write-Host "FAIL: function saw $inner, expected 2"; exit 1 }
if ($x -ne 1) { Write-Host "FAIL: caller's x became $x, expected 1"; exit 1 }
Write-Host 'PASS'
exit 0
