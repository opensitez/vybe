# vybe-test: powershell/ast_parsing/ast_parsing_binary_expression_ast
$sb = { $x + 5 }
$binAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.BinaryExpressionAst] }, $true)
if ($binAst.Operator.ToString() -ne "Plus") {
    Write-Host "FAIL: BinaryExpressionAst operator expected Plus, got $($binAst.Operator)"
    exit 1
}
Write-Host "PASS"
exit 0
