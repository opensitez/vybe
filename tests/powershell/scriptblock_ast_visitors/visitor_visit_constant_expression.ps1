# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_constant_expression
$sb = { "ConstantString" }
$constants = [System.Collections.Generic.List[object]]::new()
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.ConstantExpressionAst]) {
        $constants.Add($ast.Value)
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if ($constants -notcontains "ConstantString") {
    Write-Host "FAIL: Visit ConstantExpressionAst expected 'ConstantString'"
    exit 1
}
Write-Host "PASS"
exit 0
