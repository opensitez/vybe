# vybe-test: powershell/ast_switch_statement_ast/ast_switch_clause_item2_statement_block
# The action block executed on pattern match is stored in Clause.Item2 as a StatementBlockAst
$code = @"
switch (`$action) {
    'start' {
        `$status = 'running'
        Log-Status `$status
    }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$switchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.SwitchStatementAst] }, $true)
$clause = $switchAst.Clauses[0]

$actionBlock = $clause.Item2

if ($actionBlock -isnot [System.Management.Automation.Language.StatementBlockAst]) {
    Write-Host "FAIL: expected StatementBlockAst for clause action, got '$($actionBlock.GetType().Name)'"
    exit 1
}

if ($actionBlock.Statements.Count -ne 2) {
    Write-Host "FAIL: expected 2 statements in clause action block, got $($actionBlock.Statements.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
