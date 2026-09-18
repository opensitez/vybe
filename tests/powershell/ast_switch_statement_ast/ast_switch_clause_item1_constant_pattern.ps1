# vybe-test: powershell/ast_switch_statement_ast/ast_switch_clause_item1_constant_pattern
# The pattern expression of each clause is stored in Clause.Item1
$code = "switch (`$name) { 'Administrator' { `$true } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
$clause = $switchAst.Clauses[0]

$patternExpr = $clause.Item1

if ($patternExpr -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
    Write-Host "FAIL: expected StringConstantExpressionAst, got '$($patternExpr.GetType().Name)'"
    exit 1
}

if ($patternExpr.Value -ne "Administrator") {
    Write-Host "FAIL: pattern value mismatch: '$($patternExpr.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
