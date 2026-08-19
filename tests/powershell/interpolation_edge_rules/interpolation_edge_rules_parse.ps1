# vybe-test: powershell/interpolation_edge_rules/parse
$template = '"$($a + $b)"'
$a = 2
$b = 3
if ((Invoke-Expression $template) -ne '5') {
    Write-Host "FAIL: parse/eval of interpolation template failed"
    exit 1
}
Write-Host 'PASS'
exit 0
