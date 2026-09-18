# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_type_double_quoted
# Double-quoted strings without variables evaluate StringConstantType to DoubleQuoted
$code = '$str = "hello double quoted world"'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.StringConstantType -ne [System.Management.Automation.Language.StringConstantType]::DoubleQuoted) {
    Write-Host "FAIL: expected DoubleQuoted, got '$($strAst.StringConstantType)'"
    exit 1
}

Write-Host "PASS"
exit 0
