# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_binary_expression
$sb = { 10 + 20 }
$foundBinary = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.BinaryExpressionAst]) {
        $script:foundBinary = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundBinary) {
    Write-Host "FAIL: Visit BinaryExpressionAst expected binary expression node"
    exit 1
}
Write-Host "PASS"
exit 0
