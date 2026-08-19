# vybe-test: powershell/comment_syntax_suite/comment_with_hash_hashhash
$joined = 1..4 -join '##'
if ($joined -ne '1##2##3##4') {
    Write-Host "FAIL: expected 1##2##3##4, got '$joined'"
    exit 1
}

$literal = 'a##b##c'
if ($literal -ne 'a##b##c') {
    Write-Host "FAIL: expected literal hash text preserved"
    exit 1
}

Write-Host 'PASS'
exit 0
