# vybe-test: powershell/whitespace_and_line_rules/padded_pipe_syntax
$result = 1, 2, 3 |
   Where-Object   { $_   -gt   1 }   |
   Select-Object  -First   1

if ($result -ne 2) {
    Write-Host "FAIL: padded pipe syntax changed pipeline output: $result"
    exit 1
}

Write-Host 'PASS'
exit 0
