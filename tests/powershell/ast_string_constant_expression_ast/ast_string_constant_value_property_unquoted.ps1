# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_value_property_unquoted
# StringConstantExpressionAst.Value extracts the payload without enclosing quotes
$code = "`$title = 'PowerShell Core Architecture'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.Value -ne "PowerShell Core Architecture") {
    Write-Host "FAIL: unquoted value mismatch: '$($strAst.Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
