# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_in_try_catch_block
$executed = $false
try {
    $true && ($executed = $true)
} catch {}
if (-not $executed) {
    Write-Host "FAIL: && inside try block failed"
    exit 1
}
Write-Host "PASS"
exit 0
