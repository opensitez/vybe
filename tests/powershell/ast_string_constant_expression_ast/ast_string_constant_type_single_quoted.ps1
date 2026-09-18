# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_type_single_quoted
# Single-quoted string literals evaluate StringConstantType to SingleQuoted
$code = "`$str = 'hello single quoted world'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.StringConstantType -ne [System.Management.Automation.Language.StringConstantType]::SingleQuoted) {
    Write-Host "FAIL: expected SingleQuoted, got '$($strAst.StringConstantType)'"
    exit 1
}

Write-Host "PASS"
exit 0
