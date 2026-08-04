# vybe-test: powershell/literals/here_string_literal
$here = @"
line1
line2
"@
if ($here -notlike '*line1*') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
