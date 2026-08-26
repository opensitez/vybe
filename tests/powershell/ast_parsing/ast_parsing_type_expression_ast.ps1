# vybe-test: powershell/ast_parsing/ast_parsing_type_expression_ast
$ast = [System.Management.Automation.Language.Parser]::ParseInput('[int]$x', [ref]$null, [ref]$null)
$typeAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TypeConstraintAst] -or $args[0] -is [System.Management.Automation.Language.TypeExpressionAst] }, $true)
if ($typeAst -eq $null) {
    Write-Host "FAIL: Type expression AST lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
