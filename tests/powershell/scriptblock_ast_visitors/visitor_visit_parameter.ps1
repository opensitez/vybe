# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_parameter
$sb = { param([string]$TargetParam) }
$foundParam = $false
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.ParameterAst]) {
        $script:foundParam = $true
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if (-not $foundParam) {
    Write-Host "FAIL: Visit ParameterAst expected parameter node"
    exit 1
}
Write-Host "PASS"
exit 0
