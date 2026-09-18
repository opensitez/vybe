# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_static_type_is_string
# StringConstantExpressionAst.StaticType evaluates to System.String
$code = "`$test = 'constant string'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.StaticType -ne [string]) {
    Write-Host "FAIL: expected StaticType [string], got '$($strAst.StaticType)'"
    exit 1
}

Write-Host "PASS"
exit 0
