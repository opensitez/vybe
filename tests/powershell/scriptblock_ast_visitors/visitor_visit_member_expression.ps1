# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_member_expression
$sb = { $obj.Property }
$foundMember = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.MemberExpressionAst]) {
        $script:foundMember = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundMember) {
    Write-Host "FAIL: Visit MemberExpressionAst expected member expression node"
    exit 1
}
Write-Host "PASS"
exit 0
