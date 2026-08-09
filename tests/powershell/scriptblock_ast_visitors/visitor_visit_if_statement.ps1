# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_if_statement
$sb = { if ($true) { 1 } }
$foundIf = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.IfStatementAst]) {
        $script:foundIf = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundIf) {
    Write-Host "FAIL: Visit IfStatementAst expected if statement node"
    exit 1
}
Write-Host "PASS"
exit 0
