# vybe-test: powershell/null_conditional/null_conditional_expression_use
$val = $null
$res = (${val}?.Length) -eq $null
if (-not $res) {
    Write-Host "FAIL: (${val}?.Length) -eq \$null expected true"
    exit 1
}
Write-Host "PASS"
exit 0
