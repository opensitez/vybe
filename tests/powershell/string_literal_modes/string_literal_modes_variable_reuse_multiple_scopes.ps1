# vybe-test: powershell/string_literal_modes/variable_reuse_multiple_scopes
$token = 'outer'
function test-scope-string-literals {
    $token = 'inner'
    return $token
}
$fromFunction = test-scope-string-literals
if ($fromFunction -ne 'inner') {
    Write-Host "FAIL: function scope interpolation returned '$fromFunction'"
    exit 1
}
if ($token -ne 'outer') {
    Write-Host "FAIL: outer scope variable was unexpectedly changed to '$token'"
    exit 1
}

Write-Host 'PASS'
exit 0
