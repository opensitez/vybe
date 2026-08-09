# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_variable_expression
$sb = { $x = 100 }
$varNames = [System.Collections.Generic.List[string]]::new()
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.VariableExpressionAst]) {
        $varNames.Add($ast.VariablePath.UserPath)
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if ($varNames -notcontains "x") {
    Write-Host "FAIL: Visit VariableExpressionAst expected variable 'x'"
    exit 1
}
Write-Host "PASS"
exit 0
