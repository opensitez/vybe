# vybe-test: powershell/comment_syntax_suite/comment_inside_array_literal
$values = @(1, # inline comment
            2,
            3)

if (($values.Count -ne 3) -or ($values[1] -ne 2)) {
    Write-Host 'FAIL: comment inside array literal broke array parsing'
    exit 1
}

Write-Host 'PASS'
exit 0
