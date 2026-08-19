# vybe-test: powershell/token_and_grammar_system/punctuation_dot_access
$obj = [pscustomobject]@{ Value = 9 }
if ($obj.Value -ne 9) {
    Write-Host "FAIL: dot-member access failed got $($obj.Value)"
    exit 1
}

Write-Host 'PASS'
exit 0
