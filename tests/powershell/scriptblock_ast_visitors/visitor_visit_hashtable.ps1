# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_hashtable
$sb = { @{ K = "V" } }
$foundHash = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.HashtableAst]) {
        $script:foundHash = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundHash) {
    Write-Host "FAIL: Visit HashtableAst expected hashtable node"
    exit 1
}
Write-Host "PASS"
exit 0
