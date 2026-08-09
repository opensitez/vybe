# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_attribute
$sb = { param([Parameter(Mandatory=$true)][string]$Text) }
$foundAttr = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.AttributeAst]) {
        $script:foundAttr = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundAttr) {
    Write-Host "FAIL: Visit AttributeAst expected attribute node"
    exit 1
}
Write-Host "PASS"
exit 0
