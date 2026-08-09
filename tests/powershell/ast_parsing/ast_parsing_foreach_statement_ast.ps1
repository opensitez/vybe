# vybe-test: powershell/ast_parsing/ast_parsing_foreach_statement_ast
$sb = { foreach ($item in $list) { $item } }
$feAst = $sb.Ast.Find({ param($ast) $ast -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)
if ($feAst.Variable.VariablePath.UserPath -ne "item") {
    Write-Host "FAIL: ForEachStatementAst variable expected 'item', got '$($feAst.Variable.VariablePath.UserPath)'"
    exit 1
}
Write-Host "PASS"
exit 0
