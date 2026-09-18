# vybe-test: powershell/ast_switch_statement_ast/ast_switch_break_inside_clause_action
# A break statement inside a switch clause action block is properly located in the AST
$code = "switch (`$num) { 1 { Write-Host 'one'; break } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
$clauseAction = $switchAst.Clauses[0].Item2

$breakNode = $clauseAction.Find({ $args[0] -is [System.Management.Automation.Language.BreakStatementAst] }, $true)

if ($breakNode -eq $null) {
    Write-Host "FAIL: BreakStatementAst not found inside clause action block"
    exit 1
}

Write-Host "PASS"
exit 0
