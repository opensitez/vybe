# vybe-test: powershell/readonly_variables/readonly_variable_subexpression
New-Variable -Name "RO_SUB" -Value 77 -Option ReadOnly
$msg = "Value is $( $RO_SUB )"
if ($msg -ne "Value is 77") {
    Write-Host "FAIL: ReadOnly variable subexpression expected 'Value is 77', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
