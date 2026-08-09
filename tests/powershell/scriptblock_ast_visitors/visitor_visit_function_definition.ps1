# vybe-test: powershell/scriptblock_ast_visitors/visitor_visit_function_definition
$sb = { function Test-Fn { "Fn" } }
$funcNames = [System.Collections.Generic.List[string]]::new()
$sb.Ast.Visit({
    param($ast)
    if ($ast -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
        $funcNames.Add($ast.Name)
    }
    return [System.Management.Automation.Language.AstVisitAction]::Continue
})
if ($funcNames -notcontains "Test-Fn") {
    Write-Host "FAIL: Visit FunctionDefinitionAst expected function 'Test-Fn'"
    exit 1
}
Write-Host "PASS"
exit 0
