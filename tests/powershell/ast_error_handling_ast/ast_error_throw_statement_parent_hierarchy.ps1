# vybe-test: powershell/ast_error_handling_ast/ast_error_throw_statement_parent_hierarchy
# ThrowStatementAst.Parent points to the containing StatementBlockAst
$code = "if (`$failed) { throw 'Operation failed' }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$throwAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ThrowStatementAst] }, $true)

$parentBlock = $throwAst.Parent

if ($parentBlock -isnot [System.Management.Automation.Language.StatementBlockAst]) {
    Write-Host "FAIL: expected parent to be StatementBlockAst, got '$($parentBlock.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
