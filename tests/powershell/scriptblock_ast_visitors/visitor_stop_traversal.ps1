# vybe-test: powershell/scriptblock_ast_visitors/visitor_stop_traversal
$sb = { $x = 1; $y = 2; $z = 3 }
$visitedCount = 0
$sb.Ast.Visit({
    param($ast)
    $script:visitedCount++
    if ($script:visitedCount -eq 1) {
        return [System.Management.Automation.Language.AstVisitAction]::Stop
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if ($visitedCount -ne 1) {
    Write-Host "FAIL: AstVisitAction::Stop expected exactly 1 node visited, got $visitedCount"
    exit 1
}
Write-Host "PASS"
exit 0
