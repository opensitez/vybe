# vybe-test: powershell/interpolation/statement_subexpression
# `$( … )` takes a STATEMENT, not only an expression: the branch that runs
# supplies the value.
$text = "v=$(if ($true) { 'yes' } else { 'no' })"
if ($text -ne 'v=yes') {
    Write-Host "FAIL: got [$text]"
    exit 1
}
Write-Host 'PASS'
exit 0
