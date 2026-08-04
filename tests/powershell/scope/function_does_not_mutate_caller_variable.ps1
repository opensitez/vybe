# vybe-test: powershell/scope/function_does_not_mutate_caller_variable
$x = 10
function TryMutate { $x = 99 }
TryMutate
if ($x -ne 10) {
    Write-Host "FAIL: function should not mutate caller's variable, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
