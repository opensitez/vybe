# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_foreach_statement
$sb = { foreach ($i in $list) { $i } }
$foundForEach = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.ForEachStatementAst]) {
        $script:foundForEach = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundForEach) {
    Write-Host "FAIL: Visit ForEachStatementAst expected foreach statement node"
    exit 1
}
Write-Host "PASS"
exit 0
