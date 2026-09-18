# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_value_escaped_double_quotes
# Backtick-escaped double quotes inside double-quoted strings are unescaped in .Value
$code = '$msg = "He said `"Hello`" to me"'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.Value -ne 'He said "Hello" to me') {
    Write-Host "FAIL: unescaped double quote value mismatch: '$($strAst.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
