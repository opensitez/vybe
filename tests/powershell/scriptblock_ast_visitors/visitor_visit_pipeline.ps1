# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_pipeline
$sb = { 1..5 | Where-Object { $_ -gt 2 } }
$foundPipe = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.PipelineAst]) {
        $script:foundPipe = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundPipe) {
    Write-Host "FAIL: Visit PipelineAst expected pipeline node"
    exit 1
}
Write-Host "PASS"
exit 0
