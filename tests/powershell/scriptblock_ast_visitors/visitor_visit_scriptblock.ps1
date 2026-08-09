# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_scriptblock
$sb = { { "Inner" } }
$foundNestedSb = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) {
        $script:foundNestedSb = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundNestedSb) {
    Write-Host "FAIL: Visit ScriptBlockExpressionAst expected nested scriptblock node"
    exit 1
}
Write-Host "PASS"
exit 0
