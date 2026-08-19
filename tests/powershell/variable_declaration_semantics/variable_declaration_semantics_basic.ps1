# vybe-test: powershell/variable_declaration_semantics/basic
if (Get-Variable -Name t_var_qa -Scope Local -ErrorAction SilentlyContinue) {
    Remove-Variable -Name t_var_qa -Scope Local -Force
}

New-Variable -Name t_var_qa -Value 123 -Option ReadOnly -Scope Local
$value = Get-Variable -Name t_var_qa -ValueOnly

if ($value -ne 123) {
    Write-Host "FAIL: variable declaration failed, got $value"
    Remove-Variable -Name t_var_qa -Scope Local -Force
    exit 1
}

Remove-Variable -Name t_var_qa -Scope Local -Force
Write-Host 'PASS'
exit 0
