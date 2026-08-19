# vybe-test: powershell/comment_syntax_suite/comment_between_then_else
$path = 'init'
if ($true) {
    $path = 'then'
} else {
    $path = 'else' # else branch marker comment
}

if ($path -ne 'then') {
    Write-Host "FAIL: comment between branches altered control flow: $path"
    exit 1
}

Write-Host 'PASS'
exit 0
