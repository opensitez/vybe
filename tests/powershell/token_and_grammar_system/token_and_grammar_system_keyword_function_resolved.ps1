# vybe-test: powershell/token_and_grammar_system/keyword_function_resolved
function token_and_grammar_system_keyword_function_resolved {
    return 17
}

if ((token_and_grammar_system_keyword_function_resolved) -ne 17) {
    Write-Host 'FAIL: function keyword/classic syntax not resolved'
    exit 1
}

Write-Host 'PASS'
exit 0
