# vybe-test: powershell/ast_parsing/ast_parsing_unary_expression_ast
$sb = { -not $flag }
$unAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.UnaryExpressionAst] }, $true)
if ($unAst.TokenKind.ToString() -ne "Not") {
    Write-Host "FAIL: UnaryExpressionAst TokenKind expected Not, got $($unAst.TokenKind)"
    exit 1
}
Write-Host "PASS"
exit 0
