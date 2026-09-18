# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_empty_double_quoted
# Empty double-quoted string "" evaluates Value to empty string and StringConstantType to DoubleQuoted
$code = '$empty = ""'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.Value -ne "") {
    Write-Host "FAIL: expected empty string Value, got '$($strAst.Value)'"
    exit 1
}

if ($strAst.StringConstantType -ne [System.Management.Automation.Language.StringConstantType]::DoubleQuoted) {
    Write-Host "FAIL: expected DoubleQuoted for empty double quotes, got '$($strAst.StringConstantType)'"
    exit 1
}

Write-Host "PASS"
exit 0
