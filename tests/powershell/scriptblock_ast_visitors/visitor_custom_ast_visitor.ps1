# vybe-test: powershell/scriptblock_ast_visitors/visitor_custom_ast_visitor
class CustomVarVisitor : System.Management.Automation.Language.AstVisitor {
    [System.Collections.Generic.List[string]]$Vars = [System.Collections.Generic.List[string]]::new()
    [System.Management.Automation.Language.AstVisitAction] VisitVariableExpression([System.Management.Automation.Language.VariableExpressionAst]$ast) {
        $this.Vars.Add($ast.VariablePath.UserPath)
        return [System.Management.Automation.Language.AstVisitAction]::Continue
    }
}
$sb = { $a = 1; $b = $a + 2 }
$visitor = [CustomVarVisitor]::new()
$sb.Ast.Visit($visitor)
if ($visitor.Vars.Count -lt 2) {
    Write-Host "FAIL: CustomVarVisitor expected to collect at least 2 variables"
    exit 1
}
Write-Host "PASS"
exit 0
