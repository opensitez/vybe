# vybe-test: powershell/ast_expandable_string_expression_ast/ast_expandable_string_multiple_variables_nested
# "User: $user from $domain" contains 2 distinct VariableExpressionAst entries in NestedExpressions
$code = '$str = "User: $user from $domain"'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$expAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)

if ($expAst.NestedExpressions.Count -ne 2) {
    Write-Host "FAIL: expected 2 nested expressions, got $($expAst.NestedExpressions.Count)"
    exit 1
}

$first = $expAst.NestedExpressions[0]
$second = $expAst.NestedExpressions[1]

if ($first.VariablePath.UserPath -ne "user" -or $second.VariablePath.UserPath -ne "domain") {
    Write-Host "FAIL: variable names mismatch: '$($first.VariablePath.UserPath)', '$($second.VariablePath.UserPath)'"
    exit 1
}

Write-Host "PASS"
exit 0
