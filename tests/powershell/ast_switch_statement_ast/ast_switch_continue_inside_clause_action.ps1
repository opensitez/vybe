# vybe-test: powershell/ast_switch_statement_ast/ast_switch_continue_inside_clause_action
# A continue statement inside a switch clause action block is properly located in the AST
$code = "switch (`$num) { 1 { continue } }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
$clauseAction = $switchAst.Clauses[0].Item2

$continueNode = $clauseAction.Find({ $args[0] -is [System.Management.Automation.Language.ContinueStatementAst] }, $true)

if ($continueNode -eq $null) {
    Write-Host "FAIL: ContinueStatementAst not found inside clause action block"
    exit 1
}

Write-Host "PASS"
exit 0
