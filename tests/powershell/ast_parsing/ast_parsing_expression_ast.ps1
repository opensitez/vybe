# vybe-test: powershell/ast_parsing/ast_parsing_expression_ast
$sb = { 42 }
$expr = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.ConstantExpressionAst] }, $true)
if ($expr.Value -ne 42) {
    Write-Host "FAIL: ConstantExpressionAst Value expected 42, got $($expr.Value)"
    exit 1
}
Write-Host "PASS"
exit 0
