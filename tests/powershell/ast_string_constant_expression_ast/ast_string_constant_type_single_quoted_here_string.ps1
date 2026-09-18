# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_type_single_quoted_here_string
# Single-quoted here-strings parse as StringConstantType.SingleQuotedHereString
$code = "`$here = @'`nLine 1`nLine 2`n'@"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.StringConstantType -ne [System.Management.Automation.Language.StringConstantType]::SingleQuotedHereString) {
    Write-Host "FAIL: expected SingleQuotedHereString, got '$($strAst.StringConstantType)'"
    exit 1
}

Write-Host "PASS"
exit 0
