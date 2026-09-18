# vybe-test: powershell/ast_switch_statement_ast/ast_switch_clause_item1_scriptblock_condition
# Dynamic scriptblock conditions in a clause are parsed as ScriptBlockExpressionAst
$code = "switch (`$file) { { `$_.Length -gt 1MB } { 'Large file' } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
$clause = $switchAst.Clauses[0]

$condExpr = $clause.Item1

if ($condExpr -isnot [System.Management.Automation.Language.ScriptBlockExpressionAst]) {
    Write-Host "FAIL: expected ScriptBlockExpressionAst for clause condition, got '$($condExpr.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
