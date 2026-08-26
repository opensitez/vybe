# vybe-test: powershell/dynamic_pipeline_variable_splatting/splatting_ast_variable_reference_validation
$ast = [System.Management.Automation.Language.Parser]::ParseInput('
Get-Process @params
', [ref]$null, [ref]$null)
$varAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)
if ($varAst -eq $null -or -not $varAst.Splatted) {
    Write-Host "FAIL: Splatted AST inspection check failed"
    exit 1
}
Write-Host "PASS"
exit 0
