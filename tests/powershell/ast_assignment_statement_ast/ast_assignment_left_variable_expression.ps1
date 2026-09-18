# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_left_variable_expression
# When assigning to a variable, AssignmentStatementAst.Left is a VariableExpressionAst
$code = "`$assignedVariable = 123"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Left -isnot [System.Management.Automation.Language.VariableExpressionAst]) {
    Write-Host "FAIL: Left was not VariableExpressionAst"
    exit 1
}

$varName = $assignAst.Left.VariablePath.UserPath

if ($varName -ne "assignedVariable") {
    Write-Host "FAIL: variable name mismatch: '$varName'"
    exit 1
}

Write-Host "PASS"
exit 0
