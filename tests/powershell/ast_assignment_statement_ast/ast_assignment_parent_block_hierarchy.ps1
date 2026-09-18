# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_parent_block_hierarchy
# AssignmentStatementAst.Parent points to the containing block AST
$code = "if (`$true) { `$assigned = 1 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

$parentBlock = $assignAst.Parent

if ($parentBlock -isnot [System.Management.Automation.Language.StatementBlockAst]) {
    Write-Host "FAIL: expected parent to be StatementBlockAst, got '$($parentBlock.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
