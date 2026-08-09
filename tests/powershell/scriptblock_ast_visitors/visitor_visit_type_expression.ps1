# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_type_expression
$sb = { [int]"50" }
$foundTypeExpr = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.TypeExpressionAst]) {
        $script:foundTypeExpr = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundTypeExpr) {
    Write-Host "FAIL: Visit TypeExpressionAst expected type expression node"
    exit 1
}
Write-Host "PASS"
exit 0
