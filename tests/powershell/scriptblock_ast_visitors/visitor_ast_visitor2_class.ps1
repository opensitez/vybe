# vybe-test: powershell/scriptblock_ast_visitors/visitor_ast_visitor2_class
class NodeCounterVisitor : System.Management.Automation.Language.AstVisitor2 {
    [int]$Count = 0
    [object] VisitStatementBlock([System.Management.Automation.Language.StatementBlockAst]$ast) {
        $this.Count++
        return [System.Management.Automation.Language.AstVisitAction]::Continue
    }
}
$sb = { $a = 1; $b = 2 }
$v = [NodeCounterVisitor]::new()
$sb.Ast.Visit($v)
if ($v.Count -lt 1) {
    Write-Host "FAIL: AstVisitor2 subclass VisitStatementBlock expected Count >= 1"
    exit 1
}
Write-Host "PASS"
exit 0
