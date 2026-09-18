# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_extent_includes_quotes_double
# StringConstantExpressionAst.Extent.Text encloses the double quotes
$code = '$raw = "double quoted"'

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.Extent.Text -ne '"double quoted"') {
    Write-Host "FAIL: extent text mismatch, expected '""double quoted""', got '$($strAst.Extent.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
