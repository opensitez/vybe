# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_array_literal
$sb = { 10, 20, 30 }
$foundArrayLit = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.ArrayLiteralAst]) {
        $script:foundArrayLit = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundArrayLit) {
    Write-Host "FAIL: Visit ArrayLiteralAst expected array literal node"
    exit 1
}
Write-Host "PASS"
exit 0
