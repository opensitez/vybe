# vybe-test: powershell/ast_string_constant_expression_ast/ast_string_constant_extent_includes_quotes_single
# StringConstantExpressionAst.Extent.Text encloses the single quotes
$code = "`$raw = 'quoted value'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$strAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true)

if ($strAst.Extent.Text -ne "'quoted value'") {
    Write-Host "FAIL: extent text mismatch, expected ''quoted value'', got '$($strAst.Extent.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
