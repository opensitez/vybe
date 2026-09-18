# vybe-test: powershell/ast_expandable_string_expression_ast/ast_expandable_string_single_variable_nested
# "Hello $name" parses as ExpandableStringExpressionAst with 1 VariableExpressionAst in NestedExpressions
$code = '$msg = "Hello $name"'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$expAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)

if ($expAst.NestedExpressions.Count -ne 1) {
    Write-Host "FAIL: expected 1 nested expression, got $($expAst.NestedExpressions.Count)"
    exit 1
}

$nested = $expAst.NestedExpressions[0]

if ($nested -isnot [System.Management.Automation.Language.VariableExpressionAst]) {
    Write-Host "FAIL: expected VariableExpressionAst, got '$($nested.GetType().Name)'"
    exit 1
}

if ($nested.VariablePath.UserPath -ne "name") {
    Write-Host "FAIL: nested variable name mismatch: '$($nested.VariablePath.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
