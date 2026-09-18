# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_value_escaped_single_quotes
# Doubled single quotes inside single-quoted strings are unescaped in .Value
$code = "`$str = 'It''s working as expected'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.Value -ne "It's working as expected") {
    Write-Host "FAIL: unescaped single quote value mismatch: '$($strAst.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
