# vybe-test: powershell/ast_loop_statements/ast_loop_foreach_variable_expression
# ForEachStatementAst.Variable represents the iteration variable
$code = "foreach (`$record in `$dataSet) { `$record.Process() }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$foreachAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)

if ($foreachAst.Variable -isnot [System.Management.Automation.Language.VariableExpressionAst]) {
    Write-Host "FAIL: Variable was not VariableExpressionAst"
    exit 1
}

$varName = $foreachAst.Variable.VariablePath.UserPath

if ($varName -ne "record") {
    Write-Host "FAIL: iteration variable mismatch, expected 'record', got '$varName'"
    exit 1
}

Write-Host "PASS"
exit 0
