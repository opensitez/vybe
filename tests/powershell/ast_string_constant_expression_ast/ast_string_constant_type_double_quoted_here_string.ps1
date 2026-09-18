# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_type_double_quoted_here_string
# Double-quoted here-strings without variables parse as DoubleQuotedHereString
$code = "`$hereDouble = @`"`nMulti line double`nno variables`n`"@"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.StringConstantType -ne [System.Management.Automation.Language.StringConstantType]::DoubleQuotedHereString) {
    Write-Host "FAIL: expected DoubleQuotedHereString, got '$($strAst.StringConstantType)'"
    exit 1
}

Write-Host "PASS"
exit 0
