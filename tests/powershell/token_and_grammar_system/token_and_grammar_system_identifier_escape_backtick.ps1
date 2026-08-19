# vybe-test: powershell/token_and_grammar_system/identifier_escape_backtick
${na me} = 7
if (${na me} -ne 7) {
    Write-Host "FAIL: escaped identifier with backtick-space-like syntax not honored, got ${na me}"
    exit 1
}

Write-Host 'PASS'
exit 0
