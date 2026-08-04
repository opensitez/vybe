# vybe-test: powershell/variable_expansion/expand_variable_in_heredoc
$name = 'Vybe'
$text = @"
Hello $name
"@
if ($text -notlike '*Hello Vybe*') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
