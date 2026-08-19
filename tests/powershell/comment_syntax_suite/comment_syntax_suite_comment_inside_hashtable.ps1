# vybe-test: powershell/comment_syntax_suite/comment_inside_hashtable
$hash = @{
    Name = 'x' # comment inside hashtable key
    Value = 7
}

if ($hash.Value -ne 7 -or $hash.Name -ne 'x') {
    Write-Host 'FAIL: comment inside hashtable changed structure'
    exit 1
}

Write-Host 'PASS'
exit 0
