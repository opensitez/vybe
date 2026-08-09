# vybe-test: powershell/ast_parsing/ast_parsing_find_all_functions
$sb = { function Test-A { 1 }; function Test-B { 2 } }
$funcs = $sb.Ast.FindAll({ param($ast) $ast -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
if ($funcs.Count -ne 2) {
    Write-Host "FAIL: FunctionDefinitionAst count expected 2, got $($funcs.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
