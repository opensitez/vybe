# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_try_statement
$sb = { try { 1 } catch { 2 } }
$foundTry = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.TryStatementAst]) {
        $script:foundTry = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundTry) {
    Write-Host "FAIL: Visit TryStatementAst expected try statement node"
    exit 1
}
Write-Host "PASS"
exit 0
