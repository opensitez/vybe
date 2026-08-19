# vybe-test: powershell/token_and_grammar_system/comment_freeze
$tokenText = '#not-a-comment'
$label = "token=$tokenText"
$payload = @($label)

if ($payload.Count -ne 1) {
    Write-Host "FAIL: unexpected payload count $($payload.Count)"
    exit 1
}

if ($label -ne 'token=#not-a-comment') {
    Write-Host "FAIL: comment marker should remain literal, got '$label'"
    exit 1
}

Write-Host 'PASS'
exit 0
