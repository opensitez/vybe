# vybe-test: powershell/whitespace_and_line_rules/mixed_spacing_in_array_literal
$vals = @(1, 2 , 3 ,  4)
if (($vals.Count -ne 4) -or ($vals[0] -ne 1) -or ($vals[3] -ne 4)) {
    Write-Host 'FAIL: mixed spacing in array literal altered parse/elements'
    exit 1
}

Write-Host 'PASS'
exit 0
