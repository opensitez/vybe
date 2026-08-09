# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_command
$sb = { Get-Process -Name "pwsh" }
$cmdNames = [System.Collections.Generic.List[string]]::new()
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.CommandAst]) {
        $cmdNames.Add($ast.GetCommandName())
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if ($cmdNames -notcontains "Get-Process") {
    Write-Host "FAIL: Visit CommandAst expected command 'Get-Process'"
    exit 1
}
Write-Host "PASS"
exit 0
