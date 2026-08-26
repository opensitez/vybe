# vybe-test: powershell/token_and_grammar_system/token_and_grammar_system_comment_block_nested_not_supported
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
