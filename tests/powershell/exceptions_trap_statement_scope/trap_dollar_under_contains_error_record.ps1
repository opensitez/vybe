# vybe-test: powershell/exceptions_trap_statement_scope/trap_dollar_under_contains_error_record
$capturedMsg = ""
function Test-TrapDollarUnder {
    trap {
        $script:capturedMsg = $_.Exception.Message
        continue
    }
    throw "TrapMsgCheck"
}
Test-TrapDollarUnder
if ($capturedMsg -ne "TrapMsgCheck") {
    Write-Host "FAIL: `$_ in trap block failed, got '$capturedMsg'"
    exit 1
}
Write-Host "PASS"
exit 0
