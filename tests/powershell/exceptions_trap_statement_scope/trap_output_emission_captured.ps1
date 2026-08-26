# vybe-test: powershell/exceptions_trap_statement_scope/trap_output_emission_captured
function Test-TrapOutput {
    trap {
        "TRAP_OUTPUT"
        continue
    }
    1 / 0
    "NORMAL_OUTPUT"
}
$res = @(Test-TrapOutput)
if ($res -notcontains "TRAP_OUTPUT" -or $res -notcontains "NORMAL_OUTPUT") {
    Write-Host "FAIL: Trap output emission failed, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
