# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_left_index_expression
# When assigning to an indexed collection element, Left is an IndexExpressionAst
$code = "`$items[0] = 'first_element'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Left -isnot [System.Management.Automation.Language.IndexExpressionAst]) {
    Write-Host "FAIL: Left was not IndexExpressionAst"
    exit 1
}

$targetCollection = $assignAst.Left.Target.VariablePath.UserPath
if ($targetCollection -ne "items") {
    Write-Host "FAIL: target collection name mismatch: '$targetCollection'"
    exit 1
}

Write-Host "PASS"
exit 0
