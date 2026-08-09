# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_unary_expression
$sb = { -not $flag }
$foundUnary = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.UnaryExpressionAst]) {
        $script:foundUnary = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundUnary) {
    Write-Host "FAIL: Visit UnaryExpressionAst expected unary expression node"
    exit 1
}
Write-Host "PASS"
exit 0
