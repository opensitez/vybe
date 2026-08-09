# vybe-test: powershell/psvariable_objects/psvariable_scope_parameter
Set-Variable -Name "GlobalScopedVar" -Value "GlobalVal" -Scope Global
if ($global:GlobalScopedVar -ne "GlobalVal") {
    Write-Host "FAIL: Set-Variable -Scope Global expected \$global:GlobalScopedVar='GlobalVal'"
    exit 1
}
Write-Host "PASS"
exit 0
